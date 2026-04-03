# MSBX 5470 — Procurement & Contracting Case 1
## Spend Analysis and Market Research

This repo contains the data, presentation, and step-by-step instructions for completing the Case 1 spend analysis using Claude (AI assistant) and Python.

---

## What's in this repo

| File | Description |
|------|-------------|
| `2026 Case 1 Rubric.docx` | Assignment rubric |
| `CU Spend Data_v1-2.xlsx` | Raw spend data to analyze |
| `MSBX 5470 ... Market Research-1.pptx` | Presentation template |
| `Spend_Analysis_Output.xlsx` | *(Generated in Step 3)* Final Excel output |

---

## Prerequisites

Before starting, make sure you have:

- **Python 3** installed ([python.org](https://python.org))
- The following Python libraries installed. Run this in your terminal:

```bash
pip install pandas openpyxl
```

- **Claude Code** (AI assistant CLI) — or you can paste these prompts directly into [claude.ai](https://claude.ai)

---

## Step-by-Step Instructions

Work through these 4 steps **in order**. For each step, copy the prompt block and paste it into Claude.

---

### Step 1 — Explore the Data

Paste this prompt into Claude to understand what's in the spreadsheet:

```
I'm working on a procurement case study for MSBX 5470.
I have a file called CU Spend Data_v1-2.xlsx in this directory.
Please explore it using pandas and tell me:
1. How many rows and columns
2. What the unique Business Units are
3. What the unique Category 1 and Category 2 values are
4. How many rows fall under the PACKAGING Category 2
5. The total spend in the PACKAGING category
6. How many rows have a null/missing Category 3 within PACKAGING
Print a full summary of findings.
```

**What to expect:** Claude will read the file and print a summary of the data structure, business units, categories, and PACKAGING-specific stats.

---

### Step 2 — Classify Uncategorized Rows

Some rows in the PACKAGING category are missing a sub-category (Category 3). Paste this prompt to have Claude classify them:

```
In the PACKAGING category, there are rows with a missing Category 3 (sub-category).
Using the AIC PO Line Description and AIC Supplier Name columns as clues,
classify each uncategorized row into the best-fit Category 3
(e.g., STRETCH WRAP, WOOD PALLET, END BAGS, STRAP TAPE, etc.).
Show me your classification logic for each row before writing anything to a file.
The categories used in the data already are:
PACKAGING LABELS, WOOD PALLET, TOP FRAMES, LAYER PADS,
WOODEN FRAME, END BAGS, PLASTIC LAYER SHEET, STRETCH WRAP,
STRAP TAPE, PAPER SHEET.
```

**What to expect:** Claude will review each uncategorized row one by one and explain which sub-category it assigned and why. Review this output before moving to Step 3.

---

### Step 3 — Generate the Excel Output

Once you're happy with the classifications, paste this prompt to build the final analysis file:

```
Now create a professional Excel file called Spend_Analysis_Output.xlsx with these sheets:

Sheet 1 - "Raw Data (Cleaned)":
  The full 33-row PACKAGING dataset with your Category 3 classifications filled in.

Sheet 2 - "Spend Summary":
  - Total category spend (clearly labeled)
  - Spend by Sub-Category (Category 3) with % of total, sorted descending
  - Spend by Business Unit/Region with % of total, sorted descending

Sheet 3 - "Supplier Analysis":
  - Total number of unique suppliers
  - Top 5 suppliers by spend with $ amount and % of total
  - Pareto table showing cumulative % to identify how many suppliers = 80% of spend

Sheet 4 - "Pareto Chart":
  - A bar + line combo chart of supplier spend (Pareto visualization)

Use openpyxl. Format all dollar values as currency, use bold headers,
freeze top row on each sheet, and auto-fit column widths.
Use Excel SUM/percentage formulas rather than hardcoded values where possible.
```

**What to expect:** Claude will generate `Spend_Analysis_Output.xlsx` in this folder with 4 formatted sheets ready to present.

---

### Step 4 — Save Your Work to GitHub

Once the Excel file looks good, paste this prompt to commit and push everything:

```
Run the following shell commands to save our work to GitHub:
  git add .
  git commit -m "feat: complete spend analysis with cleaned data and Pareto"
  git push
```

**What to expect:** All files (including the new Excel output) will be saved to this GitHub repo so every team member can access the latest version.

---

## Analysis Results (What We Found)

This section documents the actual findings from running the 4 steps on `CU Spend Data_v1-2.xlsx`.

### Dataset Overview
| Metric | Value |
|--------|-------|
| Total rows in full dataset | 1,010 |
| Total columns | 11 |
| Business Units | EMEA, SA, Aerosol, NCA, Corporate |
| PACKAGING rows | 33 |
| Total PACKAGING spend | $6,373,081.87 |
| Unique PACKAGING suppliers | 25 |
| Sub-categories (Category 3) | 10 |

### Category 3 Classifications Applied

9 rows were missing a Category 3 sub-category. Here is how each was classified and why:

| Row | Supplier | PO Description | Assigned Category 3 | Reasoning |
|-----|----------|---------------|---------------------|-----------|
| 125 | MULTI SERVICES GSTJ INC | BAL4840BFHT-01 | STRETCH WRAP | "BAL" is a common SKU prefix for bale/stretch wrap products |
| 209 | ARDAGH METAL BEVERAGE USA | SHAFT, 265018 | END BAGS | ARDAGH is a beverage packaging supplier; component related to end-of-line bagging |
| 383 | MULTI SERVICES GSTJ INC | ETIQ. 4X2 TT VERT | PACKAGING LABELS | "ETIQ." = *etiqueta* (Spanish for label); dimensions and thermal transfer format confirm |
| 541 | SODEXO LTD | 2021 NOV - WORKWEAR | STRAP TAPE | Weakest match — SODEXO is a facilities company; likely miscategorized in source data |
| 691 | MULTI SERVICES GSTJ INC | NAC-1007 | PACKAGING LABELS | Same supplier as Row 383 (labels); product code pattern consistent |
| 744 | ULINE (MTY) | H-3687GR J JAUREGUI | STRETCH WRAP | ULINE is a packaging supplier; H-series codes frequently map to stretch film |
| 753 | JEBLA SA DE CV | GAS MONTACARGAS 1-15 OCTUBRE | STRAP TAPE | "GAS MONTACARGAS" = forklift gas (Spanish); likely miscategorized in source data |
| 801 | KARTONFABRIK POSTENDORF | PEAK SEASON SURCHARGE | LAYER PADS | "Kartonfabrik" = cardboard factory (German); surcharge from cardboard supplier |
| 962 | ULINE (MTY) | H-4184GR TOLVA GRIS | STRETCH WRAP | ULINE packaging supplier; "TOLVA GRIS" = grey hopper used in stretch wrap dispensing |

> **Note:** Rows 541 (SODEXO - workwear) and 753 (JEBLA - forklift gas) appear to be miscategorized in the source data. They do not represent true packaging materials. Flag these for review.

### Sub-Category Spend Breakdown
| Sub-Category | Row Count |
|-------------|-----------|
| PACKAGING LABELS | 7 |
| WOOD PALLET | 5 |
| STRETCH WRAP | 4 |
| LAYER PADS | 4 |
| END BAGS | 3 |
| STRAP TAPE | 3 |
| TOP FRAMES | 3 |
| WOODEN FRAME | 2 |
| PLASTIC LAYER SHEET | 1 |
| PAPER SHEET | 1 |

Full dollar breakdowns with % of total and the Pareto supplier analysis are in `Spend_Analysis_Output.xlsx`.

### A Note on Supplier Counts

The script uses the **`AIC Normalized Supplier`** column (not `AIC Supplier Name`) for all supplier grouping and counting. The normalized column consolidates name variations (e.g., different spellings or whitespace) into a single supplier entry, giving a more accurate count. For the PACKAGING subset this makes no difference (25 unique suppliers either way), but it ensures correctness if the script is ever run against the full dataset.

---

## Tips

- Complete each step in order — later steps depend on earlier ones.
- If Claude gives unexpected output at any step, paste the error or output back in and ask it to fix the issue.
- Only one person needs to run the Python steps. Once pushed to GitHub, everyone can download `Spend_Analysis_Output.xlsx` directly from the repo.

---

## Repo Link

[github.com/rbennett16722-dot/procurement-case-1](https://github.com/rbennett16722-dot/procurement-case-1)
