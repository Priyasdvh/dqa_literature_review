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

OUT_DIM_PATH = BASE_DIR / "DQ_Dimensions_Figure_Colorful.png"
OUT_METH_PATH = BASE_DIR / "DQ_Methods_Figure_Colorful.png"


# ============================================================
# 2. Figure settings
# ============================================================

DPI = 300
SORT_BY_YEAR = True

# The DQD figure is wider because it has nine categories.
FIGSIZE_DIM = (15, 12)
FIGSIZE_METH = (10, 12)

EXPECTED_NUMBER_OF_STUDIES = 19

EXPECTED_DIM_COUNTS = np.array(
    [12, 9, 8, 6, 5, 3, 2, 1, 1],
    dtype=int,
)

EXPECTED_METH_COUNTS = np.array(
    [15, 11, 11, 3],
    dtype=int,
)


# ============================================================
# 3. Excel layout
# ============================================================

# Zero-based column index for the author/year labels.
AUTHOR_COL = 1

# Excel rows 4–22 contain the 19 included studies.
# iloc excludes the final row, so 3:22 selects rows 4–22.
DATA_START_ROW = 3
DATA_END_ROW = 22

# Excel columns D–L: DQ dimensions.
DIM_COLS = [3, 4, 5, 6, 7, 8, 9, 10, 11]

# Excel columns M–P: DQA methods.
METH_COLS = [12, 13, 14, 15]


# ============================================================
# 4. Figure labels
# ============================================================

DIM_LABELS = [
    "Completeness",
    "Conformance",
    "Plausibility",
    "Consistency",
    "Accuracy /\nCorrectness /\nValidity",
    "Concordance",
    "Currency",
    "Uniqueness",
    "Temporal\nRelationships",
]

METH_LABELS = [
    "Element-level\nAssessment",
    "Cross-database or\nreference-based\nComparison",
    "Rule-based\nChecks",
    "Structured DQA\nframework /\nProgramme",
]


# ============================================================
# 5. Figure colours
# ============================================================

DIM_COLORS = [
    "#2A9D8F",  # Completeness
    "#2F80C1",  # Conformance
    "#8E44AD",  # Plausibility
    "#C58B00",  # Consistency
    "#E03131",  # Accuracy / Correctness / Validity
    "#D81B60",  # Concordance
    "#7A7A7A",  # Currency
    "#2E8B57",  # Uniqueness
    "#4D4D4D",  # Temporal relationships
]

METH_COLORS = [
    "#3AAFA9",  # Element-level assessment
    "#5DADE2",  # Cross-database/reference comparison
    "#7E57C2",  # Rule-based checks
    "#E67E22",  # Structured framework/programme
]


# ============================================================
# 6. Load and validate the Excel data
# ============================================================

if not EXCEL_PATH.exists():
    raise FileNotFoundError(
        f"Excel file not found:\n{EXCEL_PATH}\n\n"
        "Keep DQ_Dimensions_Matrix.xlsx in the same folder "
        "as this Python script."
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

if len(authors) != EXPECTED_NUMBER_OF_STUDIES:
    raise ValueError(
        f"Expected {EXPECTED_NUMBER_OF_STUDIES} studies, "
        f"but found {len(authors)}.\n"
        "Check DATA_START_ROW and DATA_END_ROW."
    )

if any(author == "" for author in authors):
    raise ValueError(
        "At least one selected Excel row has no author/year label."
    )


# ============================================================
# 7. Convert matrix cells to binary values
# ============================================================

def to_binary(value) -> int:
    """
    Convert common positive matrix indicators to 1.

    Accepted positive values:
    ✓, x, 1, yes, true

    Blank cells and all other values become 0.
    """
    if pd.isna(value):
        return 0

    text = str(value).strip().casefold()

    return 1 if text in {"✓", "✔", "x", "1", "yes", "true"} else 0


dim_matrix = (
    data[DIM_COLS]
    .apply(lambda column: column.map(to_binary))
    .to_numpy(dtype=int)
)

meth_matrix = (
    data[METH_COLS]
    .apply(lambda column: column.map(to_binary))
    .to_numpy(dtype=int)
)


# ============================================================
# 8. Extract years and sort publications
# ============================================================

def extract_year(label: str) -> int:
    """Extract the final four-digit publication year from a label."""
    matches = re.findall(r"\b(?:19|20)\d{2}\b", str(label))

    if not matches:
        raise ValueError(
            f"Could not extract a publication year from: {label!r}"
        )

    return int(matches[-1])


if SORT_BY_YEAR:
    years = np.array(
        [extract_year(author) for author in authors],
        dtype=int,
    )

    # Descending internally means the oldest publication appears
    # at the top of the figure because smaller y values are lower.
    order = np.argsort(years, kind="stable")[::-1]

    authors = [authors[index] for index in order]
    dim_matrix = dim_matrix[order, :]
    meth_matrix = meth_matrix[order, :]


# ============================================================
# 9. Validate totals
# ============================================================

dimension_totals = dim_matrix.sum(axis=0)
method_totals = meth_matrix.sum(axis=0)

print(f"Studies loaded: {len(authors)}")

print("\nData quality dimensions:")
for label, count in zip(DIM_LABELS, dimension_totals):
    clean_label = label.replace("\n", " ")
    print(f"{clean_label:42s} {int(count):>2}")

print("\nData quality assessment methods:")
for label, count in zip(METH_LABELS, method_totals):
    clean_label = label.replace("\n", " ")
    print(f"{clean_label:42s} {int(count):>2}")


if not np.array_equal(dimension_totals, EXPECTED_DIM_COUNTS):
    raise ValueError(
        "\nDQD totals do not match the expected validated counts.\n"
        f"Expected: {EXPECTED_DIM_COUNTS.tolist()}\n"
        f"Found:    {dimension_totals.tolist()}\n\n"
        "Check the DQ-dimension marks in the Excel matrix."
    )

if not np.array_equal(method_totals, EXPECTED_METH_COUNTS):
    raise ValueError(
        "\nDQA-method totals do not match the expected validated counts.\n"
        f"Expected: {EXPECTED_METH_COUNTS.tolist()}\n"
        f"Found:    {method_totals.tolist()}\n\n"
        "Check the DQA-method marks in the Excel matrix."
    )


# ============================================================
# 10. Shared drawing function
# ============================================================

def draw_figure(
    matrix: np.ndarray,
    labels: list[str],
    title: str,
    colors: list[str],
    figsize: tuple[float, float],
    output_path: Path,
    star_size: float,
    x_label_fontsize: float,
    x_label_rotation: float,
    x_label_pad: float,
    y_label_fontsize: float = 9,
    title_fontsize: float = 16,
) -> None:
    """Draw a bar chart and study-by-category star matrix."""

    number_studies, number_columns = matrix.shape

    if number_columns != len(labels):
        raise ValueError(
            f"The matrix has {number_columns} columns, "
            f"but {len(labels)} labels were supplied."
        )

    if number_columns != len(colors):
        raise ValueError(
            f"The matrix has {number_columns} columns, "
            f"but {len(colors)} colours were supplied."
        )

    x_positions = np.arange(number_columns)
    column_totals = matrix.sum(axis=0)

    fig, (bar_axis, matrix_axis) = plt.subplots(
        nrows=2,
        ncols=1,
        sharex=True,
        figsize=figsize,
        gridspec_kw={
            "height_ratios": [2, 7],
            "hspace": 0.02,
        },
    )

    fig.patch.set_facecolor("white")

    fig.suptitle(
        title,
        fontsize=title_fontsize,
        fontweight="bold",
        y=0.985,
    )

    # --------------------------------------------------------
    # Top bar chart
    # --------------------------------------------------------

    bars = bar_axis.bar(
        x_positions,
        column_totals,
        color=colors,
        width=0.60,
        edgecolor="white",
        linewidth=0.8,
    )

    for bar, total, color in zip(
        bars,
        column_totals,
        colors,
    ):
        bar_axis.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 0.25,
            str(int(total)),
            ha="center",
            va="bottom",
            fontsize=11,
            fontweight="bold",
            color=color,
        )

    bar_axis.set_ylabel(
        "Studies (n)",
        fontsize=11,
        labelpad=8,
    )

    bar_axis.set_ylim(
        0,
        max(column_totals) + 4,
    )

    bar_axis.tick_params(
        axis="y",
        labelsize=10,
    )

    bar_axis.spines["top"].set_visible(False)
    bar_axis.spines["right"].set_visible(False)
    bar_axis.spines["left"].set_color("#CCCCCC")
    bar_axis.spines["bottom"].set_color("#CCCCCC")
    bar_axis.set_facecolor("white")

    # --------------------------------------------------------
    # Bottom star matrix
    # --------------------------------------------------------

    for row_index in range(number_studies):
        for column_index in range(number_columns):
            if matrix[row_index, column_index] == 1:
                matrix_axis.scatter(
                    x_positions[column_index],
                    row_index + 1,
                    marker="*",
                    s=star_size,
                    color=colors[column_index],
                    edgecolors="white",
                    linewidths=0.3,
                    zorder=3,
                )

    # Vertical dotted grid lines
    for x_position in x_positions:
        matrix_axis.axvline(
            x_position,
            color="#DDDDDD",
            linestyle=":",
            linewidth=0.8,
            zorder=0,
        )

    # Horizontal dotted grid lines
    for y_position in range(1, number_studies + 1):
        matrix_axis.axhline(
            y_position,
            color="#DDDDDD",
            linestyle=":",
            linewidth=0.8,
            zorder=0,
        )

    matrix_axis.set_xlim(
        -0.8,
        number_columns - 0.2,
    )

    matrix_axis.set_ylim(
        0,
        number_studies + 1,
    )

    matrix_axis.set_yticks(
        np.arange(1, number_studies + 1)
    )

    matrix_axis.set_yticklabels(
        authors,
        fontsize=y_label_fontsize,
    )

    matrix_axis.set_ylabel(
        "Publications (first author and year of publication)",
        fontsize=11,
        labelpad=10,
    )

    matrix_axis.set_xticks(x_positions)

    matrix_axis.set_xticklabels(
        labels,
        fontsize=x_label_fontsize,
        rotation=x_label_rotation,
        ha="right",
        rotation_mode="anchor",
    )

    matrix_axis.tick_params(
        axis="x",
        pad=x_label_pad,
    )

    matrix_axis.spines["top"].set_visible(False)
    matrix_axis.spines["right"].set_visible(False)
    matrix_axis.spines["left"].set_color("#CCCCCC")
    matrix_axis.spines["bottom"].set_color("#CCCCCC")
    matrix_axis.set_facecolor("white")

    # Leave additional room for the rotated x-axis labels.
    fig.tight_layout(
        rect=[0.01, 0.045, 0.995, 0.965],
        h_pad=0.4,
    )

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
# 11. Figure 1: DQ dimensions
# ============================================================

draw_figure(
    matrix=dim_matrix,
    labels=DIM_LABELS,
    title="Data Quality Dimensions Addressed",
    colors=DIM_COLORS,
    figsize=FIGSIZE_DIM,
    output_path=OUT_DIM_PATH,
    star_size=145,
    x_label_fontsize=11,
    x_label_rotation=32,
    x_label_pad=8,
    y_label_fontsize=9,
    title_fontsize=16,
)


# ============================================================
# 12. Figure 2: DQA methods
# ============================================================

draw_figure(
    matrix=meth_matrix,
    labels=METH_LABELS,
    title="Data Quality Assessment Methods Applied",
    colors=METH_COLORS,
    figsize=FIGSIZE_METH,
    output_path=OUT_METH_PATH,
    star_size=155,
    x_label_fontsize=10,
    x_label_rotation=35,
    x_label_pad=7,
    y_label_fontsize=9,
    title_fontsize=16,
)


print("\nDone.")