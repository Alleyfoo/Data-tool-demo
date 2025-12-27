# Data Tool Demo

Provider data ingestor demo that standardizes messy Excel/CSV files into a clean, validated schema. Includes a CLI/batch runner, Streamlit UI, and a lightweight Tk GUI for template creation.

## What it does
- Standardizes spreadsheet headers and datatypes using reusable templates (`.df-template.json`) with pandera validation.
- Batch processing: drop files in `data/input`, outputs land in `data/output`, successes are archived, and failures are quarantined with logs.
- Streamlit experience: Upload, Mapping, Query Builder (SQL-like), Diagnostics, Template Library, Combine & Export.
- Combine or join cleaned outputs into a master report with the `combine` command.
- Desktop-friendly Tk GUI (`python main.py gui`) for quick template creation without Streamlit.

## Repo contents (demo-ready)
- Sample input: `data/input/sample_messy.xlsx`
- Sample template: `data/schemas/sample_messy.df-template.json`
- Sample cleaned output (reference): `data/output/sample_messy_clean.xlsx`
- Docs: `QUICKSTART.md` (guided run) and `GUIDE_QUERY_BUILDER.md` (visual builder details)
- Config: `src/config.yaml` (base synonyms) and `src/config.user.yaml` (auto-learned from saved templates)

## Quickstart (about a minute)
1) Python 3.10+ recommended.
2) Create and activate a virtual env.
3) Install deps: `pip install -r requirements.txt`
4) Run the sample: `python main.py run --target-dir data/input`
   - Cleaned output appears in `data/output/sample_messy_clean.xlsx`
   - Source is archived to `data/archive/`; failures (if any) go to `data/quarantine/` with a log.

## Run the apps
- Streamlit UI: `streamlit run app.py` (pages: Dashboard, Upload, Mapping, Query Builder, Diagnostics, Template Library, Combine & Export)
- CLI batch: `python main.py run --target-dir data/input --output-fmt xlsx --validation-level coerce`
- Combine cleaned files: `python main.py combine --input-dir data/output --pattern "*.xlsx" --mode concat --output Master_Sales_Report.xlsx`
- Tk GUI: `python main.py gui`

## Working with templates
- Save templates next to source files or in `data/schemas/`.
- Template Library (Streamlit) can inspect/run `.df-template.json` files and batch process an input directory.
- Configurable synonyms live in `src/config.yaml`; user-learned header hints accumulate in `src/config.user.yaml`.

## Testing
```bash
pip install -r requirements.txt
pytest
```

## Notes for the demo repo
- Generated artifacts live in `data/output/`, `data/archive/`, and `data/quarantine/`; it's safe to clear them between runs.
- To add database pulls, install drivers and set passwords via environment variables (for example, `<NAME>_PASSWORD` for a saved connection). See `Onboarding.md` for connection tips.
