# Goodreads To-Read Ranker

A tiny local app for ranking books from a Goodreads export using pairwise comparisons and Glicko-2.

## What it does

Upload your Goodreads CSV/XLS/XLSX export.

The app automatically filters:

    Exclusive Shelf == "to-read"

You then repeatedly answer:

    Which of these two books would you rather read?

The app uses the answers to produce a ranked list.

## Requirements

Python 3.10 or newer.

## Windows

Double-click:

    run.bat

The first run creates a virtual environment and installs the dependencies.

## macOS / Linux

Run:

    ./run.sh

## Manual launch

    python -m pip install -r requirements.txt
    python -m streamlit run app.py

## Data

Everything runs locally.

The uploaded Goodreads data is not sent to a remote service.

Ranking progress is stored in:

    .ranker_state/

Do not delete that folder if you want to resume an unfinished ranking.

## Output

The exported Excel workbook contains:

### Ranking

The current ranking, with:

- Rank
- Title
- Author
- Rating
- Rating uncertainty
- Number of comparisons
- Wins
- Losses
- Ties

### Goodreads Data

The original Goodreads data plus ranking information.

### Comparisons

Every decision made during the session.

## Ranking method

The app uses Glicko-2.

Each book starts at:

    Rating: 1500
    RD: 350
    Volatility: 0.06

The application chooses comparisons adaptively, favouring books that:

- have relatively few comparisons
- have high rating uncertainty
- have similar ratings to their potential opponent

This means a list of hundreds of books does not require every possible pair to be compared.

## Resetting a ranking

Delete the corresponding file in:

    .ranker_state/

Or simply upload a changed Goodreads export. The file fingerprint changes and a new ranking session is created.

## Privacy

This is intended as a local application.

No Goodreads account login is required.
No Goodreads API is required.
No book data needs to leave the computer.