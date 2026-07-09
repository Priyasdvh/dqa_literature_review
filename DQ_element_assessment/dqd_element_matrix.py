"""
DQD -> Issue Theme -> Affected Data Element  (dot matrix figure)
================================================================
Reads the matrix workbook and renders the hierarchical dot-matrix figure:
rows are issue themes grouped by DQ dimension, columns are data elements,
and a coloured dot marks each theme->element link (coloured by dimension).

Input  : DQD_Theme_Element_Matrix.xlsx   (sheet "Matrix", with X marks)
Output : DQD_Theme_Element_Matrix.png
"""

import os
import glob
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# ── Config ────────────────────────────────────────────────────────────────────
EXCEL_PATH  = "DQD_Theme_Element_Matrix.xlsx"
SHEET       = "Matrix"
OUTPUT_PATH = "DQD_Theme_Element_Matrix.png"
DPI         = 300
DOT_SIZE    = 320          # marker size for each link (stars read better larger)
TITLE       = "Data Quality Theme & Affected Data Elements"

# The six data-element columns, in display order (must match the Excel headers)
ELEMENT_COLS = [
    "Diagnostic codes", "Demographic fields", "Temporal / date fields",
    "Procedure & service codes", "Pharmacy / drug records",
    "Patient & record linkage fields",
]

# DQ-dimension display order (top -> bottom) and one signature colour each
DQD_ORDER = [
    "Completeness", "Conformance", "Plausibility", "Consistency",
    "Accuracy / Correctness", "Concordance", "Uniqueness",
    "Currency", "Relevance/Validity", "Temporal Relationships",
]
DQD_COLORS = {
    "Completeness": "#1B9E8A", "Conformance": "#2E7FB8", "Plausibility": "#8E44AD",
    "Consistency": "#B8860B", "Accuracy / Correctness": "#D62728",
    "Concordance": "#C2185B", "Uniqueness": "#2E8B57", "Currency": "#7F7F7F",
    "Relevance/Validity": "#E07B39", "Temporal Relationships": "#555555",
}

# Two-line wrapping for long labels (purely cosmetic)
WRAP = {
    "Atemporal / temporal plausibility": "Atemporal / temporal\nplausibility",
    "Accuracy / Correctness": "Accuracy /\nCorrectness",
}


# ── Locate + load ─────────────────────────────────────────────────────────────
def _locate(path):
    if os.path.exists(path):
        return path
    hits = glob.glob(f"**/{os.path.basename(path)}", recursive=True)
    if hits:
        print(f"Found '{os.path.basename(path)}' at: {hits[0]}")
        return hits[0]
    raise FileNotFoundError(f"Could not find '{path}'.")


df = pd.read_excel(_locate(EXCEL_PATH), sheet_name=SHEET)
df.columns = [str(c).strip() for c in df.columns]

# Drop the totals row and anything without a dimension/theme
df = df[df["DQ Dimension"].notna() & df["Issue theme"].notna()].copy()
df = df[df["DQ Dimension"].astype(str).str.strip().isin(DQD_ORDER)]

# A cell counts as a link if it holds an "X" (case-insensitive)
def _is_mark(v):
    return str(v).strip().upper() == "X"

# Order rows by dimension, preserving sheet order of themes within a dimension
df["_rank"] = df["DQ Dimension"].apply(lambda d: DQD_ORDER.index(d))
df = df.sort_values(["_rank"], kind="stable").reset_index(drop=True)

records = []
for _, r in df.iterrows():
    elems = {c for c in ELEMENT_COLS if _is_mark(r.get(c, ""))}
    records.append({"dqd": r["DQ Dimension"].strip(),
                    "theme": str(r["Issue theme"]).strip(),
                    "elements": elems})

n_rows = len(records)
n_cols = len(ELEMENT_COLS)
print(f"Loaded {n_rows} themes x {n_cols} elements")


# ── Draw ──────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(11, 12))
fig.patch.set_facecolor("white")
ax.set_facecolor("white")

y_of = {i: n_rows - i for i in range(n_rows)}   # first record at the top

# Column headers (rotated)
for c, name in enumerate(ELEMENT_COLS):
    ax.text(c, n_rows + 1.1, name, ha="left", va="bottom", fontsize=10,
            color="#333333", rotation=20, rotation_mode="anchor")

# Faint grid only (no shaded bands)
for c in range(n_cols):
    ax.axvline(c, color="#E2E2E2", lw=0.8, zorder=0)
for i in range(n_rows):
    ax.axhline(y_of[i], color="#EFEFEF", lw=0.6, zorder=0)

# Dimension group spans (for the coloured brackets on the left)
bounds = {}
for i, r in enumerate(records):
    bounds.setdefault(r["dqd"], [i, i])[1] = i

# Theme labels + dots
for i, r in enumerate(records):
    y = y_of[i]
    color = DQD_COLORS.get(r["dqd"], "#444444")
    label = WRAP.get(r["theme"], r["theme"])
    ax.text(-0.7, y, label, ha="right", va="center", fontsize=10, color="#222222")
    for c, name in enumerate(ELEMENT_COLS):
        if name in r["elements"]:
            ax.scatter(c, y, s=DOT_SIZE, color=color, marker="*", zorder=3,
                       edgecolors="white", linewidths=0.5)

# Dimension brackets + labels
x_dqd = -5.4
for dqd, (a, b) in bounds.items():
    y_top, y_bot = y_of[a] + 0.45, y_of[b] - 0.45
    color = DQD_COLORS.get(dqd, "#444444")
    ax.plot([x_dqd + 0.15, x_dqd + 0.15], [y_bot, y_top],
            color=color, lw=3, solid_capstyle="round", zorder=2)
    ax.text(x_dqd, (y_top + y_bot) / 2, WRAP.get(dqd, dqd).replace(" / ", " /\n"),
            ha="right", va="center", fontsize=10, fontweight="bold", color=color)

# Frame / axes off
ax.set_xlim(-6.2, n_cols - 0.3)
ax.set_ylim(0.2, n_rows + 3.0)
ax.set_xticks([]); ax.set_yticks([])
for s in ax.spines.values():
    s.set_visible(False)

# Column-group captions
ax.text(-6.2, n_rows + 2.4, "DQ Dimension", fontsize=10, fontweight="bold",
        
        fontstyle="italic", color="#666666", ha="left")
ax.text(-0.7, n_rows + 2.4, "Issue Theme", fontsize=10, fontweight="bold",
        fontstyle="italic", color="#666666", ha="right")
ax.set_title(TITLE, fontsize=13, fontweight="bold", pad=22)

plt.tight_layout()
fig.savefig(OUTPUT_PATH, dpi=DPI, bbox_inches="tight", facecolor="white")
print(f"Saved -> {OUTPUT_PATH}")