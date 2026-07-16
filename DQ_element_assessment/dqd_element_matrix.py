"""
DQD -> Issue Theme -> Affected Data Element
============================================

Input:
    DQD_Theme_Element_Matrix.xlsx
    Sheet: Matrix

Outputs:
    Figure_5_Data_Quality_Themes_and_Affected_Elements.png
    Figure_5_Data_Quality_Themes_and_Affected_Elements.pdf
"""

import os
import glob
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

# ── Config ────────────────────────────────────────────────────────────────────
EXCEL_PATH  = "DQD_Theme_Element_Matrix.xlsx"
SHEET       = "Matrix"
OUTPUT_PATH = "Figure_5_Data_Quality_Themes_and_Affected_Elements.png"
DPI         = 300
DOT_SIZE    = 360
TITLE       = ""

# The six data-element columns, in display order (must match the Excel headers)
ELEMENT_COLS = [
    "Diagnostic codes", "Demographic fields", "Temporal / date fields",
    "Procedure & service codes", "Pharmacy / drug records",
    "Patient & record linkage fields",
]

# DQ-dimension display order (top -> bottom) and one signature colour each
DQD_ORDER = [
    "Completeness",
    "Conformance",
    "Plausibility",
    "Consistency",
    "Accuracy / Correctness / Validity",
    "Concordance",
    "Currency",
    "Uniqueness",
    "Temporal Relationships",
]

DQD_COLORS = {
    "Completeness": "#1B9E8A",
    "Conformance": "#2E7FB8",
    "Plausibility": "#8E44AD",
    "Consistency": "#B8860B",
    "Accuracy / Correctness / Validity": "#D62728",
    "Concordance": "#C2185B",
    "Currency": "#7F7F7F",
    "Uniqueness": "#2E8B57",
    "Temporal Relationships": "#555555",
}

# Two-line wrapping for long labels (purely cosmetic)
WRAP = {
    "Atemporal / temporal plausibility": "Atemporal / temporal\nplausibility",
    "Accuracy / Correctness": "Accuracy /\nCorrectness",
}

## ── Locate and load ───────────────────────────────────────────────────────────
def _locate(path):
    if os.path.exists(path):
        return path

    hits = glob.glob(
        f"**/{os.path.basename(path)}",
        recursive=True,
    )

    if hits:
        print(
            f"Found '{os.path.basename(path)}' at: "
            f"{hits[0]}"
        )
        return hits[0]

    raise FileNotFoundError(
        f"Could not find '{path}'."
    )


df = pd.read_excel(
    _locate(EXCEL_PATH),
    sheet_name=SHEET,
)

df.columns = [
    str(column).strip()
    for column in df.columns
]

# Remove the totals row and rows without a dimension or theme.
df = df[
    df["DQ Dimension"].notna()
    & df["Issue theme"].notna()
].copy()

df = df[
    df["DQ Dimension"]
    .astype(str)
    .str.strip()
    .isin(DQD_ORDER)
]


# A cell represents a link when it contains "X".
def _is_mark(value):
    return str(value).strip().upper() == "X"


# Order rows by DQ dimension while preserving the original
# order of themes within each dimension.
df["_rank"] = df["DQ Dimension"].apply(
    lambda dimension: DQD_ORDER.index(dimension)
)

df = (
    df.sort_values(
        ["_rank"],
        kind="stable",
    )
    .reset_index(drop=True)
)


records = []

for _, row in df.iterrows():
    elements = {
        column
        for column in ELEMENT_COLS
        if _is_mark(row.get(column, ""))
    }

    records.append(
        {
            "dqd": str(
                row["DQ Dimension"]
            ).strip(),
            "theme": str(
                row["Issue theme"]
            ).strip(),
            "elements": elements,
        }
    )


n_rows = len(records)
n_cols = len(ELEMENT_COLS)

print(
    f"Loaded {n_rows} themes x "
    f"{n_cols} elements"
)


# ── Draw figure ───────────────────────────────────────────────────────────────
fig, ax = plt.subplots(
    figsize=(11, 12)
)

fig.patch.set_facecolor("white")
ax.set_facecolor("white")

# Position the first record at the top.
y_of = {
    index: n_rows - index
    for index in range(n_rows)
}


# ── Column headers ────────────────────────────────────────────────────────────
for column_index, column_name in enumerate(
    ELEMENT_COLS
):
    ax.text(
        column_index,
        n_rows + 1.1,
        column_name,
        ha="left",
        va="bottom",
        fontsize=10,
        color="#333333",
        rotation=20,
        rotation_mode="anchor",
    )


# ── Faint grid lines ─────────────────────────────────────────────────────────
for column_index in range(n_cols):
    ax.axvline(
        column_index,
        color="#E2E2E2",
        linestyle="-",
        linewidth=0.8,
        zorder=0,
    )

for row_index in range(n_rows):
    ax.axhline(
        y_of[row_index],
        color="#EFEFEF",
        linestyle="-",
        linewidth=0.6,
        zorder=0,
    )


# ── Outer border around the element matrix ───────────────────────────────────
matrix_border = Rectangle(
    (-0.5, 0.5),
    width=n_cols,
    height=n_rows,
    fill=False,
    edgecolor="#808080",
    linewidth=0.8,
    zorder=2,
)

ax.add_patch(matrix_border)


# ── Determine dimension-group boundaries ─────────────────────────────────────
bounds = {}

for index, record in enumerate(records):
    bounds.setdefault(
        record["dqd"],
        [index, index],
    )[1] = index


# ── Theme labels and stars ────────────────────────────────────────────────────
for row_index, record in enumerate(records):
    y_position = y_of[row_index]

    color = DQD_COLORS.get(
        record["dqd"],
        "#444444",
    )

    theme_label = WRAP.get(
        record["theme"],
        record["theme"],
    )

    ax.text(
        -0.7,
        y_position,
        theme_label,
        ha="right",
        va="center",
        fontsize=10,
        color="#222222",
    )

    for column_index, column_name in enumerate(
        ELEMENT_COLS
    ):
        if column_name in record["elements"]:
            ax.scatter(
                column_index,
                y_position,
                s=DOT_SIZE,
                color=color,
                marker="*",
                edgecolors="white",
                linewidths=0.5,
                zorder=3,
            )


# ── Dimension brackets and labels ─────────────────────────────────────────────
x_dqd = -5.4

for dqd, (first_index, last_index) in bounds.items():
    y_top = y_of[first_index] + 0.45
    y_bottom = y_of[last_index] - 0.45

    color = DQD_COLORS.get(
        dqd,
        "#444444",
    )

    ax.plot(
        [x_dqd + 0.15, x_dqd + 0.15],
        [y_bottom, y_top],
        color=color,
        linewidth=3,
        solid_capstyle="round",
        zorder=2,
    )

    dimension_label = WRAP.get(
        dqd,
        dqd,
    ).replace(
        " / ",
        " /\n",
    )

    ax.text(
        x_dqd,
        (y_top + y_bottom) / 2,
        dimension_label,
        ha="right",
        va="center",
        fontsize=10,
        fontweight="bold",
        color=color,
    )


# ── Axis limits and appearance ────────────────────────────────────────────────
ax.set_xlim(
    -6.2,
    n_cols - 0.3,
)

ax.set_ylim(
    0.2,
    n_rows + 3.0,
)

ax.set_xticks([])
ax.set_yticks([])

# Hide the main axes frame. The custom rectangle remains visible.
for spine in ax.spines.values():
    spine.set_visible(False)


# ── Column-group captions ─────────────────────────────────────────────────────
ax.text(
    -6.2,
    n_rows + 2.4,
    "DQ Dimension",
    fontsize=10,
    fontweight="bold",
    fontstyle="italic",
    color="#666666",
    ha="left",
)

ax.text(
    -0.7,
    n_rows + 2.4,
    "Issue Theme",
    fontsize=10,
    fontweight="bold",
    fontstyle="italic",
    color="#666666",
    ha="right",
)


# Add a title only if one is supplied.
if TITLE:
    ax.set_title(
        TITLE,
        fontsize=13,
        fontweight="bold",
        pad=22,
    )


# ── Save figure ───────────────────────────────────────────────────────────────
plt.tight_layout()

fig.savefig(
    OUTPUT_PATH,
    dpi=DPI,
    bbox_inches="tight",
    pad_inches=0.05,
    facecolor="white",
)

pdf_path = (
    os.path.splitext(OUTPUT_PATH)[0]
    + ".pdf"
)

fig.savefig(
    pdf_path,
    bbox_inches="tight",
    pad_inches=0.05,
    facecolor="white",
)

plt.show()
plt.close(fig)

print(f"Saved PNG: {OUTPUT_PATH}")
print(f"Saved PDF: {pdf_path}")