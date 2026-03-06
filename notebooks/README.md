# Notebook Test Suite (Windows)

Run these notebooks on a licensed Windows machine with Microsoft PowerPoint installed.

## Prerequisites

- Python environment with project dependencies installed
- Jupyter Notebook or JupyterLab
- Run notebooks from repository root (or from `notebooks/`)

## Recommended order

1. `01_engine_session_smoke.ipynb`
2. `02_discovery_and_slide_management.ipynb`
3. `03_placeholders_background.ipynb`
4. `04_snapshot_and_reports.ipynb`

## Notes

- On a proper Windows COM host, engine should be `COM`.
- `04_snapshot_and_reports.ipynb` also runs `scripts/windows_com_smoke.py` and reads latest report output.
- Generated outputs are stored under `artifacts/notebook-tests/` and `artifacts/com-smoke/`.
