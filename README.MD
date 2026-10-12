# Goodreads To-Read Ranker

A desktop Python app for turning a Goodreads export into a ranked reading queue using human pairwise choices instead of raw popularity or static sorting rules.

This project reads a Goodreads Excel export, presents two books at a time, and asks you to choose the one you prefer, tie, or skip. It then tracks those choices with a Glicko-style rating model and gradually estimates which books belong at the top of your to-read list.

The app is built around `ranker.py` and is designed for Windows use, with a convenience launcher in `ranker.bat`.

## What it does

- Imports a Goodreads export (`.xlsx` or `.xlsm`)
- Normalizes shelf data such as `to-read`, `currently-reading`, `read`, and `ignore`
- Ranks books by adaptive human preference using repeated pairwise comparisons
- Lets you choose from two books using mouse clicks or keyboard shortcuts
- Supports two ranking scopes:
  - `ACTIVE`: currently reading + to-read
  - `READ`: all-time read books
- Tracks confidence and stability metrics for Top 10 and Top 25 outcomes
- Exports a ranked workbook that adds a `Ranking - ...` sheet and a `Summary` sheet
- Saves state in a `.ranker_state` folder beside the source workbook so progress can resume later
- Lets you update source book metadata and goals directly from the UI

## Project structure

- `ranker.py` — main app; the full GUI and ranking engine
- `ranker.bat` — Windows launcher that creates a local `.venv` when needed and starts the app
- `requirements.txt` — dependency list for the core app (`openpyxl`)
- `vis.py` — separate visualization utility for plotting ranked book relationships from an Excel workbook (not the main app)
- `archive/` — historical prototype versions of the ranker and visualizer

## Requirements

- Python 3.10+
- Windows is the primary target, but the script is also likely runnable in a standard Python desktop environment with Tk installed
- `openpyxl` (installed automatically by the batch file or via `pip install -r requirements.txt`)

## Quick start

### Option 1: Use the Windows launcher

1. Open a terminal in the project folder.
2. Run:

```bat
ranker.bat
```

This will:

- create a `.venv` under the project if it does not exist
- install `openpyxl` if needed
- launch `ranker.py`

### Option 2: Run directly with Python

```bat
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe ranker.py
```

## How to use the app

1. Click `Open Excel` in the app.
2. Select a Goodreads export workbook.
3. The app reads the workbook and builds a library of books from the `Title` column.
4. It presents a pair of books for comparison.
5. Use:
   - `1` or left arrow to pick the left book
   - `2` or right arrow to pick the right book
   - `↑` to mark a tie
   - `4` or `S` to skip
   - `U` to undo
6. The ranking engine updates its internal ratings and narrows the queue toward better evidence.
7. When you are ready, use `Ctrl+E` or the export button to save a ranked workbook.

## App behavior and flow

The program is intentionally opinionated about data sources:

- It expects a Goodreads export with at least a `Title` column.
- It reads fields such as `Title`, `Author l-f`, `Exclusive Shelf`, `ISBN`, `Goals`, and metadata used for book details.
- It stores internal progress in a hidden `.ranker_state` folder next to the source workbook.
- It preserves your original workbook, creates an exported ranking workbook, and writes the ranked results into a new sheet.

The ranking engine uses a lightweight adaptive selection system:

- Each book gets a Glicko-like rating and uncertainty value.
- Pair comparisons update the rating score.
- The app chooses comparisons that are informative rather than random noise.
- It tries to surface the books that are most uncertain near the Top 25 boundary, while also respecting active reading status.

## Keyboard shortcuts

| Key | Action |
|---|---|
| `1` / `Left` | Choose left book |
| `2` / `Right` | Choose right book |
| `↑` | Mark tie |
| `4` / `S` | Skip comparison |
| `U` | Undo previous choice |
| `I` | Mark left book as Ignore |
| `R` | Mark left book as Read |
| `C` | Mark left book as Currently Reading |
| `Shift + I/R/C` | Same actions for the right book |
| `Ctrl + O` | Open Goodreads Excel |
| `Ctrl + E` | Export ranking |

## Exported workbook output

When you export, the app creates a workbook based on the source file name plus `_ranked.xlsx` (unless a different output path is chosen).

The exported workbook contains:

- a `Ranking - ACTIVE` or `Ranking - READ` sheet
- a `Summary` sheet with metadata and confidence estimates
- updated book metadata and goals written back to the original Goodreads data sheet

The ranking spreadsheet includes columns such as:

- Display
- Statistical Rank
- Title
- Author
- Status
- Rating
- Rating Uncertainty
- Rank Low / Rank High
- Top 10 % / Top 25 %
- Comparisons, Wins, Losses, Ties

## Notes on `vis.py`

`vis.py` is not required for the main ranking workflow. It is a separate data-visualization script that loads workbook sheets, normalizes titles/authors, and builds a Plotly graph from ranking data.

It is best understood as a custom visualization helper for a specific workbook schema rather than a general-purpose app entry point. If you want to use it, you may need to install extra dependencies such as `pandas` and `plotly` manually.

## `archive/`

The `archive/` directory contains earlier versions of the ranker and visualizer. These are useful for:

- comparing historical iterations
- tracing behavior changes across versions
- debugging compatibility issues with older exports

## Troubleshooting

- If the app says the workbook does not look like a Goodreads export, make sure the file includes a `Title` column.
- If no books load, verify the workbook is a valid Excel export and not a CSV or a non-Books data sheet.
- If a saved ranking state seems stale, delete the `.ranker_state` folder next to the workbook and reopen the file.
- If the app fails to open because of missing dependencies, reinstall the requirements using the project venv or `pip install -r requirements.txt`.

## License

There is no license file in this repository, so the project should be treated as source code without an explicit license declaration unless otherwise specified by the repository owner.

## Summary

This repository is a personal book-ranking tool for making better reading-priority decisions from Goodreads data. It is intentionally interactive, local-first, and focused on human judgment rather than automated sorting.
