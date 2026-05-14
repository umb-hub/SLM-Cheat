"""Shared utilities for SLM Anti-Cheat bibliometric analysis."""

from __future__ import annotations

import re
import textwrap
from itertools import combinations
from pathlib import Path

import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────
_ROOT = Path(__file__).resolve().parent.parent

PATHS = {
    "raw": _ROOT / "data" / "raw",
    "processed": _ROOT / "data" / "processed",
    "outputs": _ROOT / "outputs",
    "scripts": _ROOT / "scripts",
}

RAW_EXCEL = PATHS["raw"] / "export.xlsx"
SHEET_DATA = "Screened"
SHEET_CB = "Codebook"
SHEET_PRISMA = "PRISMA Log"

# ── Style palette ─────────────────────────────────────────────────────────────
STYLE = {
    "relevance_level": {
        "core": "#1565C0",
        "relevant": "#F57C00",
        "peripheral": "#9E9E9E",
        "borderline": "#C62828",
    },
    "contribution_type": {
        "technical": "#1976D2",
        "socio_behavioral": "#F57C00",
        "conceptual_taxonomic": "#388E3C",
        "legal_governance": "#7B1FA2",
        "review_survey": "#00838F",
        "peripheral": "#9E9E9E",
        "other_unclear": "#BDBDBD",
    },
    "study_approach": {
        "detection": "#1565C0",
        "prevention": "#2E7D32",
        "conceptualization": "#F9A825",
        "behavioral_explanation": "#E64A19",
        "infrastructure_support": "#00695C",
        "governance": "#6A1B9A",
        "measurement": "#0277BD",
        "offensive": "#C62828",
        "not_applicable": "#BDBDBD",
    },
    "phase": {
        "Phase 1 (2000–12)": "#90CAF9",
        "Phase 2 (2012–19)": "#FFB74D",
        "Phase 3 (2019–26)": "#A5D6A7",
    },
    "continent": {
        "Asia": "#E53935",
        "North America": "#1E88E5",
        "Europe": "#43A047",
        "South America": "#FB8C00",
        "Oceania": "#8E24AA",
        "Africa": "#F4511E",
        "Middle East": "#00ACC1",
        "Unknown": "#9E9E9E",
    },
}

# ── Country → continent mapping ───────────────────────────────────────────────
_CONTINENT_MAP: dict[str, str] = {
    # Asia
    "China": "Asia",
    "South Korea": "Asia",
    "Japan": "Asia",
    "India": "Asia",
    "Taiwan": "Asia",
    "Singapore": "Asia",
    "Hong Kong": "Asia",
    "Malaysia": "Asia",
    "Thailand": "Asia",
    "Indonesia": "Asia",
    "Pakistan": "Asia",
    "Bangladesh": "Asia",
    "Vietnam": "Asia",
    "Philippines": "Asia",
    "Sri Lanka": "Asia",
    "Nepal": "Asia",
    "Kazakhstan": "Asia",
    "Uzbekistan": "Asia",
    "Myanmar": "Asia",
    "Cambodia": "Asia",
    "Macau": "Asia",
    # Middle East
    "Iran": "Middle East",
    "Turkey": "Middle East",
    "Saudi Arabia": "Middle East",
    "Israel": "Middle East",
    "United Arab Emirates": "Middle East",
    "Jordan": "Middle East",
    "Egypt": "Middle East",
    "Qatar": "Middle East",
    "Kuwait": "Middle East",
    "Oman": "Middle East",
    "Lebanon": "Middle East",
    "Iraq": "Middle East",
    # Europe
    "United Kingdom": "Europe",
    "Germany": "Europe",
    "France": "Europe",
    "Spain": "Europe",
    "Italy": "Europe",
    "Netherlands": "Europe",
    "Sweden": "Europe",
    "Switzerland": "Europe",
    "Norway": "Europe",
    "Finland": "Europe",
    "Denmark": "Europe",
    "Belgium": "Europe",
    "Austria": "Europe",
    "Poland": "Europe",
    "Czech Republic": "Europe",
    "Portugal": "Europe",
    "Greece": "Europe",
    "Hungary": "Europe",
    "Romania": "Europe",
    "Russia": "Europe",
    "Ukraine": "Europe",
    "Serbia": "Europe",
    "Croatia": "Europe",
    "Slovakia": "Europe",
    "Slovenia": "Europe",
    "Bulgaria": "Europe",
    "Estonia": "Europe",
    "Latvia": "Europe",
    "Lithuania": "Europe",
    "Ireland": "Europe",
    "Luxembourg": "Europe",
    "Iceland": "Europe",
    "Malta": "Europe",
    # North America
    "United States": "North America",
    "Canada": "North America",
    "Mexico": "North America",
    # South America
    "Brazil": "South America",
    "Argentina": "South America",
    "Chile": "South America",
    "Colombia": "South America",
    "Peru": "South America",
    "Venezuela": "South America",
    "Ecuador": "South America",
    "Uruguay": "South America",
    "Bolivia": "South America",
    "Paraguay": "South America",
    # Oceania
    "Australia": "Oceania",
    "New Zealand": "Oceania",
    # Africa
    "South Africa": "Africa",
    "Nigeria": "Africa",
    "Kenya": "Africa",
    "Morocco": "Africa",
    "Tunisia": "Africa",
    "Algeria": "Africa",
    "Ethiopia": "Africa",
    "Ghana": "Africa",
}


def assign_continent(country: str) -> str:
    """Map a country name to its continent."""
    if not isinstance(country, str):
        return "Unknown"
    return _CONTINENT_MAP.get(country.strip(), "Unknown")


# ── Keyword synonym map ───────────────────────────────────────────────────────
_KW_SYNONYMS: dict[str, str] = {
    "cheats": "cheating",
    "cheat": "cheating",
    "online game": "online games",
    "online gaming": "online games",
    "peer to peer": "peer-to-peer",
    "p2p": "peer-to-peer",
    "mmog": "mmorpg",
    "massively multiplayer online game": "mmorpg",
    "massively multiplayer online games": "mmorpg",
    "mmorpgs": "mmorpg",
    "deep learning": "deep learning",
    "machine learning": "machine learning",
    "bot": "bots",
    "fps game": "fps",
    "first-person shooter": "fps",
    "first person shooter": "fps",
    "cheat detection": "cheating detection",
    "anti cheat": "anti-cheat",
    "anticheat": "anti-cheat",
    "anti-cheating": "anti-cheat",
    "aimbots": "aimbot",
    "wallhacks": "wallhack",
    "game hacking": "game cheating",
}


def normalize_keywords(series: pd.Series) -> pd.Series:
    """Lowercase, strip, and apply synonym normalization to a keyword series."""
    normalized = series.str.lower().str.strip()
    return normalized.map(lambda kw: _KW_SYNONYMS.get(kw, kw) if isinstance(kw, str) else kw)


# ── Data I/O ──────────────────────────────────────────────────────────────────

def load_raw() -> pd.DataFrame:
    """Load the full unfiltered Screened sheet from export.xlsx."""
    return pd.read_excel(RAW_EXCEL, sheet_name=SHEET_DATA)


def load_processed(stem: str) -> pd.DataFrame:
    """Load a parquet from data/processed/ by stem name."""
    path = PATHS["processed"] / f"{stem}.parquet"
    return pd.read_parquet(path)


# ── Analytical helpers ────────────────────────────────────────────────────────

def split_multivalued(df: pd.DataFrame, col: str, sep: str = ";") -> pd.DataFrame:
    """Explode a multi-valued column; returns DataFrame[Record_ID, value]."""
    exploded = (
        df[["Record_ID", col]]
        .dropna(subset=[col])
        .assign(**{col: lambda x: x[col].str.split(sep)})
        .explode(col)
    )
    exploded[col] = exploded[col].str.strip()
    exploded = exploded[exploded[col] != ""]
    return exploded.rename(columns={col: "value"}).reset_index(drop=True)


def freq_table(series: pd.Series, label: str = "value") -> pd.DataFrame:
    """Return DataFrame[value, N, pct] sorted descending. pct = 100*N/len(series)."""
    total = len(series)
    counts = series.value_counts(dropna=False)
    df = pd.DataFrame({label: counts.index, "N": counts.values})
    df["pct"] = (df["N"] / total * 100).round(1)
    return df.reset_index(drop=True)


def crosstab_with_pct(
    df: pd.DataFrame, row: str, col: str
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return (raw crosstab, row-normalized % crosstab)."""
    raw = pd.crosstab(df[row], df[col], margins=True)
    pct = pd.crosstab(df[row], df[col], normalize="index").mul(100).round(1)
    return raw, pct


def assign_phase(year_series: pd.Series) -> pd.Series:
    """Map year values to named study phases."""
    return pd.cut(
        year_series,
        bins=[1999, 2012, 2019, 2027],
        labels=["Phase 1 (2000–12)", "Phase 2 (2012–19)", "Phase 3 (2019–26)"],
        right=True,
    ).astype(str)


def build_cooccurrence_matrix(
    kw_lists: list[list[str]], min_cooc: int = 2
) -> dict[tuple[str, str], int]:
    """Build a keyword co-occurrence dict from a list of keyword lists."""
    counts: dict[tuple[str, str], int] = {}
    for kws in kw_lists:
        unique_kws = list(set(kws))
        for a, b in combinations(sorted(unique_kws), 2):
            key = (a, b)
            counts[key] = counts.get(key, 0) + 1
    return {k: v for k, v in counts.items() if v >= min_cooc}


def truncate_label(text: str, max_len: int = 50) -> str:
    """Truncate a string for display purposes."""
    if not isinstance(text, str):
        return str(text)
    return textwrap.shorten(text, width=max_len, placeholder="…")


def normalize_source_title(title: str) -> str:
    """Strip edition/year noise from proceedings titles for venue deduplication.

    Removes leading 4-digit years, ordinal numbers (1st/2nd/…), trailing
    venue-abbreviation+year suffixes (e.g. ', NetGames \'06'), and any
    remaining 4-digit years, so multiple editions of the same event collapse
    to a single canonical label.
    """
    if not isinstance(title, str):
        return str(title)
    s = title
    s = re.sub(r"^\d{4}\s+", "", s)                              # leading "2010 9th Annual…"
    s = re.sub(r"\b\d+(?:st|nd|rd|th)\b\s*", "", s, flags=re.IGNORECASE)  # ordinals
    s = re.sub(r",\s*\w+\s*[''']?\s*\d{2,4}\s*$", "", s)        # trailing ", NetGames '06"
    s = re.sub(r"\b(?:19|20)\d{2}\b", "", s)                    # remaining 4-digit years
    s = re.sub(r"\s{2,}", " ", s).strip().rstrip(",").strip()
    return s


# ── Figure export ─────────────────────────────────────────────────────────────

def save_figure(fig, name: str, subdir: str = "") -> None:
    """Save a Plotly figure as SVG, PNG (via kaleido) and HTML to outputs/."""
    out_dir = PATHS["outputs"] / subdir if subdir else PATHS["outputs"]
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = name.removesuffix(".png").removesuffix(".svg").removesuffix(".html")
    fig.write_image(str(out_dir / f"{stem}.svg"))
    fig.write_image(str(out_dir / f"{stem}.png"), width=1400, height=700, scale=2)
    fig.write_html(str(out_dir / f"{stem}.html"))
