import re
import webbrowser
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go


# ============================================================
# SETTINGS
# ============================================================

EXCEL_FILE = "E:\\Sorted\\BOOX\\_SECOND_BRAIN\\NOTES_LISTS\\BOOKS.xlsx"
OUTPUT_FILE = "book_tech_tree.html"

X_SPACING = 6
Y_SPACING = 5

# Node colours
CURRENTLY_READING_COLOR = "#EF4444"   # Red
DEFAULT_BOOK_COLOR = "#2563EB"        # Blue

# Multi-goal connector / junction styling
CONNECTION_COLOR = "#64748B"
JUNCTION_COLOR = "#F8FAFC"
JUNCTION_LINE_COLOR = "#38BDF8"


# ============================================================
# TEXT NORMALISATION
# ============================================================

def clean_text(value):
    """Normalise text for matching."""

    if pd.isna(value):
        return ""

    value = str(value).lower().strip()

    # Remove punctuation
    value = re.sub(r"[^\w\s]", " ", value)

    # Collapse whitespace
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def clean_author(value):
    """Normalise author names, including 'Last, First'."""

    value = clean_text(value)

    parts = value.split()

    if len(parts) >= 2:
        # Sort words so minor ordering differences don't matter.
        parts = sorted(parts)

    return " ".join(parts)


# ============================================================
# LOAD EXCEL
# ============================================================

print("Loading Excel file...")

books = pd.read_excel(
    EXCEL_FILE,
    sheet_name="BOOKS_all"
)

ranking = pd.read_excel(
    EXCEL_FILE,
    sheet_name="Ranking"
)

print(f"BOOKS_all rows: {len(books):,}")
print(f"Ranking rows:   {len(ranking):,}")


# ============================================================
# CHECK REQUIRED COLUMNS
# ============================================================

required_books = [
    "Title",
    "Author l-f",
    "Goals",
]

required_ranking = [
    "Title",
    "Author",
    "Statistical Rank",
    "Display"
]

missing_books = [
    c for c in required_books
    if c not in books.columns
]

missing_ranking = [
    c for c in required_ranking
    if c not in ranking.columns
]

if missing_books:
    raise ValueError(
        f"Missing columns in BOOKS_all: {missing_books}"
    )

if missing_ranking:
    raise ValueError(
        f"Missing columns in Ranking: {missing_ranking}"
    )


# ============================================================
# CREATE MATCHING KEYS
# ============================================================

books["title_key"] = (
    books["Title"]
    .apply(clean_text)
)

books["author_key"] = (
    books["Author l-f"]
    .apply(clean_author)
)

ranking["title_key"] = (
    ranking["Title"]
    .apply(clean_text)
)

ranking["author_key"] = (
    ranking["Author"]
    .apply(clean_author)
)


# ============================================================
# CLEAN STATISTICAL RANK
# ============================================================

ranking["Rank"] = pd.to_numeric(
    ranking["Statistical Rank"],
    errors="coerce"
)

print(
    f"Ranking rows with numeric Statistical Rank: "
    f"{ranking['Rank'].notna().sum():,}"
)


# ============================================================
# REMOVE INVALID RANKING ROWS
# ============================================================

ranking_valid = ranking[
    ranking["Rank"].notna() &
    ranking["title_key"].ne("")
].copy()


# ============================================================
# MATCH 1:
# TITLE + AUTHOR
# ============================================================

rank_by_author = (
    ranking_valid
    .drop_duplicates(
        subset=["title_key", "author_key"]
    )
    [
        [
            "title_key",
            "author_key",
            "Rank",
            "Display"
        ]
    ]
)

books = books.merge(
    rank_by_author,
    on=[
        "title_key",
        "author_key"
    ],
    how="left"
)

# Clean Display value brought across from Ranking
books["Display"] = (
    books["Display"]
    .fillna("")
    .astype(str)
    .str.strip()
)

matched_author = books["Rank"].notna()

print(
    f"Matched by Title + Author: "
    f"{matched_author.sum():,}"
)


# ============================================================
# MATCH 2:
# UNIQUE TITLE FALLBACK
# ============================================================

unique_title_ranks = (
    ranking_valid
    .groupby("title_key")
    .filter(
        lambda x: len(x) == 1
    )
    [
        [
            "title_key",
            "Rank",
            "Display"
        ]
    ]
    .drop_duplicates("title_key")
)

unmatched = books["Rank"].isna()

fallback = books.loc[
    unmatched,
    ["title_key"]
].merge(
    unique_title_ranks,
    on="title_key",
    how="left"
)

fallback.index = books.index[unmatched]

books.loc[
    fallback.index,
    "Rank"
] = fallback["Rank"]

books.loc[
    fallback.index,
    "Display"
] = fallback["Display"]

matched_total = books["Rank"].notna()

print(
    f"Matched after unique-title fallback: "
    f"{matched_total.sum():,}"
)

print(
    f"Still unmatched: "
    f"{(~matched_total).sum():,}"
)


# ============================================================
# SHOW EXAMPLES OF UNMATCHED BOOKS
# ============================================================

unmatched_books = books[
    books["Rank"].isna()
]

if len(unmatched_books):

    print(
        "\nFirst 20 unmatched books:"
    )

    print(
        unmatched_books[
            [
                "Title",
                "Author l-f"
            ]
        ]
        .head(20)
        .to_string(index=False)
    )


# ============================================================
# CLEAN BOOK FIELDS
# ============================================================

books["Title"] = (
    books["Title"]
    .fillna("Untitled")
    .astype(str)
)

books["Author l-f"] = (
    books["Author l-f"]
    .fillna("Unknown Author")
    .astype(str)
)

books["Goals"] = (
    books["Goals"]
    .fillna("")
    .astype(str)
    .str.strip()
)

books["Display"] = (
    books["Display"]
    .fillna("")
    .astype(str)
    .str.strip()
)


# ============================================================
# FINAL FILTER
# ============================================================

# A book must have BOTH:
#
#   1. Statistical Rank
#   2. At least one Goal
#
# Nothing without both is rendered.

books = books[
    books["Rank"].notna() &
    books["Goals"].ne("")
].copy()

books.reset_index(drop=True, inplace=True)

print(
    f"\nBooks rendered: {len(books):,}"
)

print(
    f"Currently Reading: "
    f"{(books['Display'].str.upper() == 'CURRENTLY READING').sum():,}"
)


# ============================================================
# STOP CLEANLY IF NOTHING MATCHED
# ============================================================

if books.empty:

    print(
        "\nERROR: No books have both a Rank and a Goal."
    )

    print(
        "\nCheck the matching diagnostics above."
    )

    raise SystemExit


# ============================================================
# SPLIT MULTIPLE GOALS
# ============================================================

# Goals are separated by semicolons:
#
#     Computer Vision; Economics
#
# A book remains ONE node even when it belongs to
# multiple Goals.

goal_df = books[
    [
        "Title",
        "Author l-f",
        "Rank",
        "Goals"
    ]
].copy()

# IMPORTANT:
# The source data uses semicolons to separate Goals.
goal_df["Goal"] = (
    goal_df["Goals"]
    .str.split(", ")
)

goal_df = goal_df.explode(
    "Goal"
)

goal_df["Goal"] = (
    goal_df["Goal"]
    .str.strip()
)

# Remove empty Goals
goal_df = goal_df[
    goal_df["Goal"].ne("")
].copy()


# ============================================================
# GOALS AND RANKS
# ============================================================

goals = sorted(
    goal_df["Goal"].unique()
)

ranks = sorted(
    books["Rank"].unique()
)

print(
    f"Goals: {len(goals):,}"
)

print(
    f"Ranks: {len(ranks):,}"
)


# ============================================================
# POSITIONS
# ============================================================

# Goals run horizontally.
x_pos = {
    goal: i * X_SPACING
    for i, goal in enumerate(goals)
}

# Ranks run vertically.
# Rank 1 = top.
y_pos = {
    rank: i * Y_SPACING
    for i, rank in enumerate(ranks)
}


# ============================================================
# BOOK POSITIONS
# ============================================================

positions = {}

for i, book in books.iterrows():

    book_goals = goal_df[
        (goal_df["Title"] == book["Title"]) &
        (goal_df["Author l-f"] == book["Author l-f"]) &
        (goal_df["Rank"] == book["Rank"])
    ]["Goal"].tolist()

    xs = [
        x_pos[g]
        for g in book_goals
    ]

    # A book with multiple Goals sits between them.
    x = sum(xs) / len(xs)

    # Rank determines vertical position.
    y = y_pos[book["Rank"]]

    positions[i] = (x, y)


# ============================================================
# BOOK NODES
# ============================================================

book_x = []
book_y = []
book_text = []
book_hover = []
book_colors = []

for i, book in books.iterrows():

    x, y = positions[i]

    book_x.append(x)
    book_y.append(y)

    book_text.append(
        f"<b>{book['Title']}</b><br>"
        f"<i>{book['Author l-f']}</i>"
    )

    book_goals = goal_df[
        (goal_df["Title"] == book["Title"]) &
        (goal_df["Author l-f"] == book["Author l-f"]) &
        (goal_df["Rank"] == book["Rank"])
    ]["Goal"].tolist()

    display_value = book["Display"].strip()

    # ========================================================
    # NODE COLOUR
    # ========================================================

    is_currently_reading = (
        display_value.upper() == "CURRENTLY READING"
    )

    book_colors.append(
        CURRENTLY_READING_COLOR
        if is_currently_reading
        else DEFAULT_BOOK_COLOR
    )

    book_hover.append(
        f"<b>{book['Title']}</b><br>"
        f"<i>{book['Author l-f']}</i>"
        f"<br><br>"
        f"<b>Rank:</b> {int(book['Rank'])}"
        f"<br>"
        f"<b>Goals:</b> "
        f"{', '.join(book_goals)}"
        f"<br>"
        f"<b>Display:</b> "
        f"{display_value or '—'}"
    )


# ============================================================
# GOAL LANES
# ============================================================

max_y = max(
    y_pos.values()
)

goal_x = []
goal_y = []

for goal in goals:

    x = x_pos[goal]

    goal_x.extend([
        x,
        x,
        None
    ])

    goal_y.extend([
        -Y_SPACING,
        max_y + Y_SPACING,
        None
    ])


# ============================================================
# MULTI-GOAL CONNECTIONS
# ============================================================

connection_x = []
connection_y = []

# Junction points
junction_x = []
junction_y = []
junction_hover = []

for i, book in books.iterrows():

    x, y = positions[i]

    book_goals = goal_df[
        (goal_df["Title"] == book["Title"]) &
        (goal_df["Author l-f"] == book["Author l-f"]) &
        (goal_df["Rank"] == book["Rank"])
    ]["Goal"].tolist()

    # Only create junctions for books that actually
    # belong to more than one Goal.
    if len(book_goals) > 1:

        for goal in book_goals:

            gx = x_pos[goal]

            if abs(x - gx) > 0.01:

                # Dotted horizontal connector
                connection_x.extend([
                    x,
                    gx,
                    None
                ])

                connection_y.extend([
                    y,
                    y,
                    None
                ])

                # ------------------------------------------------
                # NEW:
                # Add a visible junction marker where the dotted
                # connector reaches the Goal lane.
                # ------------------------------------------------

                junction_x.append(gx)
                junction_y.append(y)

                junction_hover.append(
                    f"<b>Shared Goal</b><br>"
                    f"{goal}<br><br>"
                    f"<b>Book:</b> {book['Title']}"
                )


# ============================================================
# CREATE FIGURE
# ============================================================

fig = go.Figure()


# ============================================================
# GOAL LANES
# ============================================================

fig.add_trace(
    go.Scatter(
        x=goal_x,
        y=goal_y,
        mode="lines",

        line=dict(
            color="#1E293B",
            width=2
        ),

        hoverinfo="none"
    )
)


# ============================================================
# MULTI-GOAL CONNECTIONS
# ============================================================

fig.add_trace(
    go.Scatter(
        x=connection_x,
        y=connection_y,

        mode="lines",

        line=dict(
            color=CONNECTION_COLOR,
            width=1.5,
            dash="dot"
        ),

        hoverinfo="none",

        name="Shared Goals"
    )
)


# ============================================================
# SHARED-GOAL JUNCTIONS
# ============================================================

# These markers sit ON TOP of the dotted connectors and
# clearly indicate that the line intentionally joins the
# Goal lane here.

fig.add_trace(
    go.Scatter(
        x=junction_x,
        y=junction_y,

        mode="markers",

        marker=dict(
            symbol="diamond",
            size=9,
            color=JUNCTION_COLOR,
            line=dict(
                color=JUNCTION_LINE_COLOR,
                width=2
            )
        ),

        hovertext=junction_hover,
        hoverinfo="text",

        name="Shared Goal Junctions"
    )
)


# ============================================================
# GOAL LABELS
# ============================================================

fig.add_trace(
    go.Scatter(
        x=[
            x_pos[g]
            for g in goals
        ],

        y=[
            -Y_SPACING
        ] * len(goals),

        mode="text",

        text=[
            f"<b>{g}</b>"
            for g in goals
        ],

        textposition="middle center",

        hoverinfo="none"
    )
)


# ============================================================
# BOOKS
# ============================================================

fig.add_trace(
    go.Scatter(
        x=book_x,
        y=book_y,

        mode="markers+text",

        text=book_text,

        textposition="middle right",

        marker=dict(
            size=18,

            # Each node gets its own colour.
            color=book_colors,

            line=dict(
                color="#FFFFFF",
                width=1
            )
        ),

        hovertext=book_hover,
        hoverinfo="text",

        name="Books"
    )
)


# ============================================================
# RANK LABELS
# ============================================================

fig.add_trace(
    go.Scatter(
        x=[
            -X_SPACING
        ] * len(ranks),

        y=[
            y_pos[r]
            for r in ranks
        ],

        mode="text",

        text=[
            f"<b>{int(r)}</b>"
            for r in ranks
        ],

        textposition="middle right",

        hoverinfo="none"
    )
)


# ============================================================
# FULL-BROWSER LAYOUT
# ============================================================

fig.update_layout(

    title=dict(
        text=(
            f"<b>BOOK TECH TREE</b>"
            f"<br>"
            f"<sup>"
            f"{len(books):,} ranked books · "
            f"{len(goals):,} goals"
            f"</sup>"
        ),

        x=0.01,
        y=0.99,

        xanchor="left",
        yanchor="top"
    ),

    autosize=True,

    margin=dict(
        l=70,
        r=20,
        t=75,
        b=20
    ),

    plot_bgcolor="#020617",
    paper_bgcolor="#020617",

    font=dict(
        color="#E2E8F0"
    ),

    showlegend=False,

    hovermode="closest",

    dragmode="pan",

    xaxis=dict(
        visible=False,
        fixedrange=False
    ),

    yaxis=dict(
        visible=False,

        # Rank 1 at top
        autorange="reversed",

        fixedrange=False
    )
)


# ============================================================
# WRITE HTML
# ============================================================

fig.write_html(
    OUTPUT_FILE,

    include_plotlyjs=True,

    full_html=True,

    config={
        "responsive": True,
        "scrollZoom": True,
        "displaylogo": False,

        "modeBarButtonsToRemove": [
            "lasso2d",
            "select2d"
        ]
    }
)


# ============================================================
# FORCE FULL BROWSER VIEWPORT
# ============================================================

html = Path(
    OUTPUT_FILE
).read_text(
    encoding="utf-8"
)

html = html.replace(
    "<body>",
    """
<body style="
    margin:0;
    padding:0;
    width:100vw;
    height:100vh;
    overflow:hidden;
">
"""
)

html = html.replace(
    "</body>",
    """
<script>

window.addEventListener(
    "resize",
    function() {

        const plot =
            document.querySelector(
                ".plotly-graph-div"
            );

        if (plot) {
            Plotly.Plots.resize(plot);
        }
    }
);

</script>

</body>
"""
)

Path(
    OUTPUT_FILE
).write_text(
    html,
    encoding="utf-8"
)


# ============================================================
# OPEN
# ============================================================

print(
    f"\\nCreated: {OUTPUT_FILE}"
)

webbrowser.open(
    Path(
        OUTPUT_FILE
    ).resolve().as_uri()
)
