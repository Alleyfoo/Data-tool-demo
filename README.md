# Data Tool Demo

A demo tool that takes messy spreadsheets (Excel or CSV files) and turns them into clean, consistent, ready-to-use data.

## The problem it solves

Real-world business data is rarely tidy. Different teams send the same information using different column names, formats, and layouts — one file calls it "Sales $", another calls it "Revenue", another calls it "amount_usd". Cleaning all of that by hand is slow and error-prone.

This tool automates the cleanup. You teach it once how a particular kind of file should look (using a reusable "template"), and after that it can clean new files of the same type automatically — checking that the data is valid, flagging anything that looks wrong, and producing a polished output file you can analyze or share.

## What you can do with it

- **Standardize spreadsheets** — Map inconsistent column names and data types to a clean, agreed-upon format using saved templates (small `.df-template.json` files that act as reusable rulebooks).
- **Process files in bulk** — Drop a folder of messy files in, get cleaned files out. Successful files are archived; problem files are set aside with a log explaining what went wrong.
- **Explore your data in a browser** — A Streamlit web app (a simple dashboard you run locally) provides pages for uploading, mapping columns, building queries (filtering and slicing data without writing code), reviewing diagnostics, managing your saved templates, and combining files.
- **Build templates with a desktop app** — A lightweight Tk-based desktop window (no browser needed) for creating templates quickly.
- **Combine cleaned files** — Stitch multiple cleaned files together into a single master report.

## Who it's for

Anyone who routinely receives spreadsheets from multiple sources — operations teams, analysts, finance, sales ops — and wants the cleaning, validating, and combining steps to happen the same way every time.

## What's included in this demo

- A sample messy input file: `data/input/sample_messy.xlsx`
- A sample template for that file: `data/schemas/sample_messy.df-template.json`
- A reference cleaned output: `data/output/sample_messy_clean.xlsx`
- Walkthrough docs: `QUICKSTART.md` (a guided run) and `GUIDE_QUERY_BUILDER.md` (how the visual query builder works)
- Configuration files: `src/config.yaml` (built-in column-name synonyms) and `src/config.user.yaml` (synonyms the tool learns as you save templates)

## Get started in about a minute

1. Use Python 3.10 or newer.
2. Create and activate a virtual environment (an isolated workspace for the project's dependencies).
3. Install the required packages: `pip install -r requirements.txt`
4. Run the sample: `python main.py run --target-dir data/input`
   - The cleaned result appears at `data/output/sample_messy_clean.xlsx`.
   - The original file is moved to `data/archive/`. Anything that failed validation is moved to `data/quarantine/` along with a log file explaining why.

## Running the apps

- **Web app (Streamlit):** `streamlit run app.py`
  Pages: Dashboard, Upload, Mapping, Query Builder, Diagnostics, Template Library, Combine & Export.
- **Batch run from the command line:**
  `python main.py run --target-dir data/input --output-fmt xlsx --validation-level coerce`
- **Combine cleaned files into one report:**
  `python main.py combine --input-dir data/output --pattern "*.xlsx" --mode concat --output Master_Sales_Report.xlsx`
- **Desktop template builder:** `python main.py gui`

## Working with templates

- Save templates alongside the source files they describe, or keep them in `data/schemas/`.
- The Template Library page in the web app lets you inspect a template, run it, or apply it to a whole folder of inputs at once.
- Built-in column-name synonyms live in `src/config.yaml`. As you save templates, the tool also remembers your own header preferences in `src/config.user.yaml` so it gets smarter over time.

## Running the tests

```bash
pip install -r requirements.txt
pytest
```

## A few notes about this demo repo

- Files the tool produces live in `data/output/`, `data/archive/`, and `data/quarantine/`. It's safe to delete the contents of these folders between runs.
- To pull data directly from a database instead of files, install the appropriate database driver and provide passwords through environment variables (for example, `<NAME>_PASSWORD` for a saved connection). See `Onboarding.md` for connection tips.
