# 📊 Excel Formula Practice Workbook — 22 Real-World Cases

A single, self-contained Excel workbook demonstrating 22 formula patterns that come up
constantly in day-to-day data analyst work — lookups, conditional aggregation, ranking,
text cleaning, date math, and multi-condition logic without helper columns.

Built on one 10-employee HR dataset, so every case is solving a realistic business
question ("what's the average salary in Finance?", "who's the top earner?", "flag
duplicate names") rather than an abstract textbook example.

---

## Why this project

I built this to demonstrate fluency with core Excel functions that show up in analyst
job descriptions and interviews — the kind of formulas used for ad-hoc reporting,
data validation, and quick lookups before a task graduates to SQL or Python. Every
sheet is a live, interactive answer, not a static screenshot.

## How it works

Each `CaseXX_...` sheet follows the same layout:

1. **A question** describing the task in plain English
2. **A highlighted (yellow) input cell** you can type into
3. **A live formula** that reads that input and pulls from the `Data` sheet
4. **A worked explanation** at the bottom breaking down exactly what the formula does

Nothing is hardcoded — every result is a real formula. Change the input, and the
answer recalculates instantly.

## Dataset

The `Data` sheet holds 10 employees × 10 fields: Name, Department, Salary, Age, City,
Joining Date, Bonus %, Experience, Rating, and Gender. Every case sheet references
this one table, so the whole workbook stays internally consistent.

## Skills demonstrated

| Category | Cases | Functions |
|---|---|---|
| Lookups | 01, 02, 09, 16 | `INDEX`/`MATCH`, array-based "list all matches", `VLOOKUP` (approximate match), two-way `INDEX`/`MATCH`/`MATCH` |
| Conditional aggregation | 03–07, 21 | `SUMIF`, `COUNTIF`, `AVERAGEIF`, `SUMIFS`, `COUNTIFS` (incl. range brackets) |
| Logic & error handling | 08, 12 | Nested `IF`, `IFERROR` + `MATCH` |
| Ranking & extremes | 10, 11, 19 | `RANK`, `LARGE`/`SMALL`, `INDEX`+`MATCH`+`MAX`/`MIN` |
| Text manipulation | 13, 14, 22 | `TEXTJOIN`, `LEFT`/`RIGHT`/`MID`/`FIND`, `PROPER`/`UPPER`/`LOWER` |
| Dates | 15 | `DATEDIF`, date subtraction (dynamic, driven by `TODAY()`) |
| Advanced / no helper columns | 17, 18, 20 | `COUNTIF` duplicate detection, `SUMPRODUCT`, manual running total |

## Example formula

Two-way lookup (Case16) — find any field for any employee by name, without a separate
formula per column:

```excel
=INDEX(Data!$A$2:$J$11, MATCH($B$7,Data!$A$2:$A$11,0), MATCH($B$8,Data!$A$1:$J$1,0))
```

The first `MATCH` finds *which row* (the employee), the second finds *which column*
(the field), and `INDEX` returns the value at that intersection.

## Repo contents

```
├── Excel_Cases_Solved.xlsx   # the workbook — start on the Index sheet
└── README.md
```

## How to explore it

1. Download `Excel_Cases_Solved.xlsx` and open it in Excel (or LibreOffice/Google Sheets)
2. Start on the **Index** sheet for a clickable list of all 22 cases
3. Pick any `CaseXX_...` tab, type a different value into the yellow input cell, and
   watch the result update live
4. Press `Ctrl` + `` ` `` (or **Formulas → Show Formulas**) to see every formula at once
