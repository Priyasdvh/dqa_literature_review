"""
Create two separate colourful figures from DQ_Dimensions_Matrix.xlsx:

1. DQ_Dimensions_Figure_Colorful.png
2. DQ_Methods_Figure_Colorful.png

Keep this script and the Excel file in the same folder.
"""

from pathlib import Path
import re

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


# ============================================================
# 1. File settings
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

EXCEL_PATH = BASE_DIR / "DQ_Dimensions_Matrix.xlsx"
SHEET_NAME = "DQ_Dimensions_Matrix"

OUT_DIM_PATH = BASE_DIR / "DQ_Dimensions_Figure.png"
OUT_METH_PATH = BASE_DIR / "DQ_Methods_Figure.png"


# ============================================================
# 2. Figure settings
# ============================================================

DPI = 300
SORT_BY_YEAR = True

FIGSIZE_DIM = (13, 12)
FIGSIZE_METH = (9, 12)


# ============================================================
# 3. Excel layout
# ============================================================

AUTHOR_COL = 1

# Excel rows 4–22 contain the 19 included studies.
DATA_START_ROW = 3
DATA_END_ROW = 22

# Excel columns D–L
DIM_COLS = [3, 4, 5, 6, 7, 8, 9, 10, 11]

# Excel columns M–P
METH_COLS = [12, 13, 14, 15]

DIM_LABELS = [
    "Completeness",
    "Conformance",
    "Plausibility",
    "Consistency",
    "Accuracy /\nCorrectness",
    "Concordance",
    "Currency",
    "Uniqueness",
    "Temporal\nrelationships",
]

METH_LABELS = [
    "Element-level\nassessment",
    "Cross-database\ncomparison",
    "Rule-based\nchecks",
    "Framework-based",
]

# Colourful palette similar to your earlier version
DIM_COLORS = [
    "#2A9D8F",  # Completeness
    "#2F80C1",  # Conformance
    "#8E44AD",  # Plausibility
    "#C58B00",  # Consistency
    "#E03131",  # Accuracy / Correctness
    "#D81B60",  # Concordance
    "#7A7A7A",  # Currency
    "#2E8B57",  # Uniqueness
    "#4D4D4D",  # Temporal relationships
]

METH_COLORS = [
    "#3AAFA9",  # Element-level assessment
    "#5DADE2",  # Cross-database comparison
    "#7E57C2",  # Rule-based checks
    "#E67E22",  # Framework-based
]


# ============================================================
# 4. Load and validate data
# ============================================================

if not EXCEL_PATH.exists():
    raise FileNotFoundError(
        f"Excel file not found:\n{EXCEL_PATH}\n\n"
        "Keep DQ_Dimensions_Matrix.xlsx in the same folder as this script."
    )

df_raw = pd.read_excel(
    EXCEL_PATH,
    sheet_name=SHEET_NAME,
    header=None,
)

required_last_col = max(DIM_COLS + METH_COLS)

if df_raw.shape[1] <= required_last_col:
    raise ValueError(
        "The Excel structure has changed.\n"
        f"Expected at least {required_last_col + 1} columns, "
        f"but found {df_raw.shape[1]}."
    )

data = df_raw.iloc[DATA_START_ROW:DATA_END_ROW].copy()
data.columns = range(data.shape[1])

authors = (
    data[AUTHOR_COL]
    .fillna("")
    .astype(str)
    .str.strip()
    .tolist()
)

if len(authors) != 19:
    raise ValueError(
        f"Expected 19 studies, but found {len(authors)}. "
        "Check DATA_START_ROW and DATA_END_ROW."
    )

if any(author == "" for author in authors):
    raise ValueError("At least one selected row has no author label.")


# ============================================================
# 5. Convert matrix cells to binary values
# ============================================================

def to_binary(value):
    text = str(value).strip().casefold()
    return 1 if text in {"✓", "1", "yes", "true", "x"} else 0


dim_matrix = (
    data[DIM_COLS]
    .map(to_binary)
    .to_numpy(dtype=int)
)

meth_matrix = (
    data[METH_COLS]
    .map(to_binary)
    .to_numpy(dtype=int)
)


# ============================================================
# 6. Sort studies by year
# ============================================================

def extract_year(label):
    matches = re.findall(r"\b(?:19|20)\d{2}\b", str(label))

    if not matches:
        raise ValueError(f"Could not extract a year from: {label!r}")

    return int(matches[-1])


if SORT_BY_YEAR:
    years = np.array([extract_year(author) for author in authors])
    order = np.argsort(years)[::-1]

    authors = [authors[index] for index in order]
    dim_matrix = dim_matrix[order, :]
    meth_matrix = meth_matrix[order, :]


# ============================================================
# 7. Print counts for verification
# ============================================================

print(f"Studies loaded: {len(authors)}")

print("\nData quality dimensions:")
for label, count in zip(DIM_LABELS, dim_matrix.sum(axis=0)):
    print(f"{label.replace(chr(10), ' '):35s} {int(count):>2}")

print("\nData quality assessment methods:")
for label, count in zip(METH_LABELS, meth_matrix.sum(axis=0)):
    print(f"{label.replace(chr(10), ' '):35s} {int(count):>2}")


# ============================================================
# 8. Shared drawing function
# ============================================================

def draw_figure(
    matrix,
    labels,
    title,
    colors,
    figsize,
    output_path,
    star_size,
):
    number_studies, number_columns = matrix.shape
    x_positions = np.arange(number_columns)
    column_totals = matrix.sum(axis=0)

    fig, (bar_axis, matrix_axis) = plt.subplots(
        nrows=2,
        ncols=1,
        sharex=True,
        gridspec_kw={"height_ratios": [2, 7]},
        figsize=figsize,
    )

    fig.patch.set_facecolor("white")
    fig.suptitle(
        title,
        fontsize=15,
        fontweight="bold",
        y=0.985,
    )

    # -------------------------
    # Top bar chart
    # -------------------------
    bars = bar_axis.bar(
        x_positions,
        column_totals,
        color=colors,
        width=0.60,
        edgecolor="white",
        linewidth=0.8,
    )

    for bar, total, color in zip(bars, column_totals, colors):
        bar_axis.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 0.25,
            str(int(total)),
            ha="center",
            va="bottom",
            fontsize=10,
            fontweight="bold",
            color=color,
        )

    bar_axis.set_ylabel("Studies (n)", fontsize=10, labelpad=8)
    bar_axis.set_ylim(0, max(column_totals) + 4)
    bar_axis.spines["top"].set_visible(False)
    bar_axis.spines["right"].set_visible(False)
    bar_axis.spines["left"].set_color("#CCCCCC")
    bar_axis.spines["bottom"].set_color("#CCCCCC")
    bar_axis.set_facecolor("white")

    # -------------------------
    # Bottom star matrix
    # -------------------------
    for row_index in range(number_studies):
        for column_index in range(number_columns):
            if matrix[row_index, column_index] == 1:
                matrix_axis.scatter(
                    x_positions[column_index],
                    row_index + 1,
                    marker="*",
                    s=star_size,
                    color=colors[column_index],
                    zorder=3,
                    linewidths=0.3,
                    edgecolors="white",
                )

    for x_position in x_positions:
        matrix_axis.axvline(
            x_position,
            color="#DDDDDD",
            linestyle=":",
            linewidth=0.8,
            zorder=0,
        )

    for y_position in range(1, number_studies + 1):
        matrix_axis.axhline(
            y_position,
            color="#DDDDDD",
            linestyle=":",
            linewidth=0.8,
            zorder=0,
        )

    matrix_axis.set_xlim(-0.8, number_columns - 0.2)
    matrix_axis.set_ylim(0, number_studies + 1)

    matrix_axis.set_yticks(np.arange(1, number_studies + 1))
    matrix_axis.set_yticklabels(authors, fontsize=9)

    matrix_axis.set_ylabel(
        "Publications (first author and year of publication)",
        fontsize=10,
        labelpad=8,
    )

    matrix_axis.set_xticks(x_positions)
    matrix_axis.set_xticklabels(
        labels,
        fontsize=9,
        rotation=35,
        ha="right",
        rotation_mode="anchor",
    )

    matrix_axis.spines["top"].set_visible(False)
    matrix_axis.spines["right"].set_visible(False)
    matrix_axis.spines["left"].set_color("#CCCCCC")
    matrix_axis.spines["bottom"].set_color("#CCCCCC")
    matrix_axis.set_facecolor("#FAFAFA")

    plt.tight_layout(h_pad=0.4, rect=[0, 0, 1, 0.965])

    fig.savefig(
        output_path,
        dpi=DPI,
        bbox_inches="tight",
        facecolor="white",
    )

    plt.show()
    plt.close(fig)

    print(f"\nSaved: {output_path}")


# ============================================================
# 9. Figure 1: DQ dimensions
# ============================================================

draw_figure(
    matrix=dim_matrix,
    labels=DIM_LABELS,
    title="Data quality dimensions addressed",
    colors=DIM_COLORS,
    figsize=FIGSIZE_DIM,
    output_path=OUT_DIM_PATH,
    star_size=140,
)


# ============================================================
# 10. Figure 2: Assessment methods
# ============================================================

draw_figure(
    matrix=meth_matrix,
    labels=METH_LABELS,
    title="Data quality assessment methods applied",
    colors=METH_COLORS,
    figsize=FIGSIZE_METH,
    output_path=OUT_METH_PATH,
    star_size=150,
)

print("\nDone.")
