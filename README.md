# Cheat SLM

Notebook-first workspace for a bibliometric study of an SLM using coded XLSX sources.

## What goes here

- `data/raw/`: original coded XLSX files
- `data/processed/`: cleaned tables and derived datasets
- `notebooks/`: exploratory analysis and reporting notebooks
- `scripts/`: reusable Python helpers for cleaning and analysis
- `outputs/`: charts, tables, and export files

## Setup

1. Install uv.
2. Sync dependencies:

```bash
uv sync
```

3. Start Jupyter Lab:

```bash
uv run jupyter lab
```

## Suggested workflow

1. Place the coded XLSX file in `data/raw/`.
2. Open the starter notebook in `notebooks/`.
3. Inspect sheet names, column names, and coding conventions.
4. Clean bibliometric fields such as year, authors, source, country, keywords, and citations.
5. Export cleaned data to `data/processed/` before building tables and plots.

## Practical visualization stack

- `plotly` for Sankey diagrams and most interactive charts
- `networkx` for graph construction and analysis
- `pyvis` for fast interactive network exploration
- `holoviews` + `bokeh` for chord diagrams and richer interactive graph views
- `datashader` for scaling dense network visualizations

This stack keeps notebook work interactive while still covering temporal network analysis, flow diagrams, and dense graphs.

## Notes

- Keep raw data untouched.
- Use notebooks for exploration and scripts for repeatable transforms.
- If the XLSX has multiple sheets or coding blocks, document the sheet-to-analysis mapping in the notebook.
