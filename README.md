# SPI Weekly Summary Generator

Turns the weekly Sensitive Price Indicator (SPI) Excel worksheet into a fully
formatted **"Brief on Weekly Inflation"** Word document, automatically, with
no manual copy-pasting of numbers into a report template.

Give it `Worksheet_DD_MM_YYYY.xlsx`, and it produces
`Summary_DD_MM_YYYY.docx` in the same layout as the original hand-built
reports (see `Summary_23_07_2026.docx` for a sample output and
`Worksheet_23.07.2026.xlsx` for a sample input).

## What it does

The generator reads the workbook and pulls out everything the report needs:

- **`Impact Combined`** sheet → combined SPI % change (week-on-week and
  year-on-year), plus group weights/impacts for Food, Non-Food, Utilities,
  and Transport.
- **`Annexure-V`** sheet → this week's % change by income quintile
  (Q1–Q5 and Combined).
- **`Sorted - MoM and YoY`** sheet → per-item price change lists used to
  write the narrative paragraphs (e.g. "prices of Wheat Flour, Onions and
  Tomatoes increased...").
- **`Appendix-A`** sheet → the full basket item count, used to work out how
  many items were flat/unchanged.

It also keeps a running `spi_history.csv` file alongside the script. Table 1
of the report ("current week vs. previous 5 weeks") needs six weeks of
quintile data, but a single workbook only contains the current week's
numbers — so the history CSV carries the other five weeks forward. Every run
appends the new week automatically.

Finally, it fills all of this into `spi_template.docx` (a Word template with
`{{PLACEHOLDER}}` tokens) and saves the finished report.

## Repo contents

| File | Purpose |
|---|---|
| `generate_spi_summary.py` | Core script, reads the worksheet, updates history, fills the template, writes the `.docx`. |
| `spi_template.docx` | Word template with `{{PLACEHOLDER}}` tokens that the script fills in. |
| `spi_history.csv` | Rolling history of weekly quintile % changes (auto-updated on every run). |
| `app.py` | Optional local web UI, drag and drop the worksheet in a browser instead of using the command line. |
| `watch_and_generate.py` | Optional folder watcher, auto regenerates the summary whenever any `.xlsx` in a folder changes. |
| `requirements.txt` | Python dependencies. |
| `Worksheet_23.07.2026.xlsx` | Sample input worksheet. |
| `Summary_23_07_2026.docx` | Sample output report. |

## Requirements

- Python 3
- `openpyxl`
- `python-docx`
- `flask` (only needed for the optional web UI, `app.py`)

Install everything with:

```bash
pip install -r requirements.txt openpyxl python-docx
```

## Usage

### Option 1 — Command line (core workflow)

```bash
python3 generate_spi_summary.py --excel Worksheet_DD_MM_YYYY.xlsx --output Summary_DD_MM_YYYY.docx
```

Optional flags:

- `--template`: path to a different `.docx` template (defaults to `spi_template.docx`)
- `--history`:  path to a different history CSV (defaults to `spi_history.csv`)

Keep `spi_history.csv` next to the script and reuse the same file every week
so the 5-week comparison table stays accurate.

### Option 2 — Local web UI

For a no-commands workflow:

```bash
python3 app.py
```

Then open `http://127.0.0.1:5050` in your browser, drag in this week's
worksheet, and the finished `.docx` downloads automatically. You can
optionally drop in a different template for a single run.

`app.py` must live in the same folder as `generate_spi_summary.py`,
`spi_template.docx`, and `spi_history.csv`.

### Option 3 — Watch a folder

To have the summary regenerate automatically whenever the worksheet is
updated (e.g. saved from Excel):

```bash
python3 watch_and_generate.py --dir "/path/to/your/SPI folder"
```

This doesn't care what the `.xlsx` file is named, it opens each changed
workbook, reads the report date from the `Impact Combined` sheet, and names
the output accordingly. Files that don't look like SPI worksheets (no
matching sheet/row) are skipped automatically. Leave it running and stop it
with `Ctrl+C`.

## Notes

- Item names in the narrative text are shortened using an editorial mapping
  (e.g. "Wheat Flour Bag" → "Wheat Flour") to match the style of the
  original reports; unmapped items fall back to their full worksheet
  description.
- Year on year **increases** are only called out by name once they pass a
  configurable threshold (`YOY_INCREASE_THRESHOLD`, default 13%), matching
  the convention of the original reports. All other change lists include
  every non-zero item.
