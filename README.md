<div align="center">

# Excel Skills — Forage Job Simulation

<p>
  <a href="https://www.theforage.com/simulations/forage/sponsored-content-excel-skills-jf8r">
    <img src="https://img.shields.io/badge/FORAGE-Excel%20Skills-176B58?style=for-the-badge&logo=target&logoColor=white" alt="Forage Excel Skills" />
  </a>
  <img src="https://img.shields.io/badge/STATUS-IN%20PROGRESS-E69A2D?style=for-the-badge" alt="Project status: in progress" />
  <img src="https://img.shields.io/badge/TOOL-Microsoft%20Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white" alt="Microsoft Excel" />
</p>

<p>
  <a href="https://github.com/Sohila-Khaled-Abbas/forage-excel-skills-job-simulation/commits/main">
    <img src="https://img.shields.io/github/last-commit/Sohila-Khaled-Abbas/forage-excel-skills-job-simulation?style=flat-square&label=last%20updated" alt="Last updated" />
  </a>
  <a href="https://github.com/Sohila-Khaled-Abbas/forage-excel-skills-job-simulation">
    <img src="https://img.shields.io/github/repo-size/Sohila-Khaled-Abbas/forage-excel-skills-job-simulation?style=flat-square&label=repository%20size" alt="Repository size" />
  </a>
  <a href="https://github.com/Sohila-Khaled-Abbas/forage-excel-skills-job-simulation/stargazers">
    <img src="https://img.shields.io/github/stars/Sohila-Khaled-Abbas/forage-excel-skills-job-simulation?style=flat-square&label=stars" alt="GitHub stars" />
  </a>
</p>

**A hands-on Excel learning portfolio documenting practical exercises, analytical decisions, workbook development, and lessons learned.**

[Simulation brief](https://www.theforage.com/simulations/forage/sponsored-content-excel-skills-jf8r) · [Repository files](https://github.com/Sohila-Khaled-Abbas/forage-excel-skills-job-simulation/tree/main) · [Learning log](LEARNING_LOG.md)

</div>

---

> [!NOTE]
> This is an independent learning portfolio created to document my participation and practice. It is **not** an official Forage repository and does not imply employment, sponsorship, or endorsement by Forage.

## Contents

- [Program overview](#program-overview)
- [Learning roadmap](#learning-roadmap)
- [Completed exercise: Conditional Formatting](#completed-exercise-conditional-formatting)
- [Completed exercise: Formulas and Functions](#completed-exercise-formulas-and-functions)
- [Current exercise: Build Charts and Dashboards](#current-exercise-build-charts-and-dashboards)
- [Repository guide](#repository-guide)
- [How I work through each exercise](#how-i-work-through-each-exercise)
- [Skills in scope](#skills-in-scope)
- [Evidence and sharing principles](#evidence-and-sharing-principles)

## Program overview

| Detail | Description |
|---|---|
| **Provider** | Forage Academy |
| **Simulation** | Excel Skills Virtual Experience Program |
| **Format** | Self-paced, practical job simulation |
| **Primary tool** | Microsoft Excel |
| **Portfolio status** | <img src="https://img.shields.io/badge/In%20progress-Working%20through%20the%20exercises-E69A2D?style=flat-square" alt="In progress" /> |
| **Official brief** | [Open the Forage simulation](https://www.theforage.com/simulations/forage/sponsored-content-excel-skills-jf8r) |

The purpose of this repository is to turn each exercise into a reviewable learning artifact. I use the workbook to practice, document the reasoning behind each Excel operation, validate the result, and capture the skills that transfer to data analysis and business intelligence work.

## Learning roadmap

The work is organized around a practical analyst workflow:

~~~mermaid
flowchart LR
    A["Explore the workbook"] --> B["Inspect data quality"]
    B --> C["Apply Excel techniques"]
    C --> D["Analyze patterns"]
    D --> E["Build visuals"]
    E --> F["Communicate findings"]
    classDef start fill:#E8F3FF,stroke:#2563EB,color:#123B66
    classDef work fill:#EAF7EF,stroke:#26935B,color:#164A2D
    classDef finish fill:#FFF3DF,stroke:#D28A22,color:#62400F
    class A start
    class B,C,D work
    class E,F finish
~~~

| Learning area | Practice focus | Documentation |
|---|---|---|
| **Workbook navigation and exploration** | Understand worksheet structure, field meanings, data types, filters, and efficient keyboard navigation. | [Excel keyboard shortcuts](docs/Excel%20Keyboard%20Shortcuts.docx) |
| **Data quality and preparation** | Locate blanks, formula errors, inconsistent values, and records that need investigation. | [Data preparation notes](learning-notes/data-preparation.md) |
| **Formulas and analysis** | Apply calculations, validate results, and explain assumptions. | [Formulas and analysis](learning-notes/formulas-and-analysis.md) |
| **Charts and dashboards** | Select appropriate visuals, define KPIs, and communicate patterns. | [Charts and dashboards](learning-notes/charts-and-dashboards.md) |
| **Automation** | Explore repeatable spreadsheet tasks and macro/VBA concepts where required. | [Automation and VBA](learning-notes/automation-and-vba.md) |
| **Reflection and communication** | Summarize what worked, what needs improvement, and what the findings mean. | [Reflection notes](learning-notes/reflection.md) |

## Completed exercise: Conditional Formatting

<details open>
<summary><strong>Task 01 — Put Conditional Formatting Into Practice</strong> · <code>Completed</code></summary>

The first exercise applied six conditional-formatting rules to the **Account Sales Data for Analysis** workbook. The completed workbook is available under `material/`, and the findings document records the data-quality observations and analytical takeaways.

**Workbooks**
- [Original practice workbook](material/01-conditional-formatting/Account%20Sales%20Data%20for%20Analysis.xlsx)
- [Completed workbook](material/01-conditional-formatting/Account_Sales_Data_Analysis_Completed.xlsx.xlsx)

**Findings:** [Task 01 — Conditional Formatting Findings](docs/01-conditional-formatting-findings.md)

### Key findings

| Finding | Result |
|---|---:|
| Account records | 40 |
| Missing cells | 5 |
| Text-based `Formula Error` markers | 2 |
| Product 2 marked `No` | 16 accounts |
| Numeric 5 YR CAGR ≥ 100% | 3 accounts |
| Negative 5 YR CAGR | 7 accounts |
| Top 10% of 2021 Unit Sales | 4 accounts |

### Formatting rules

| Requirement | Range | Format |
|---|---|---|
| Formula errors and text error markers | `A2:M41` | Purple fill, white font |
| Missing cells | `A2:M41` | Yellow fill |
| Product 2 not cross-sold | `F2:F41` | Orange fill |
| Numeric CAGR ≥ 100% | `M2:M41` | Green fill |
| Negative CAGR | `M2:M41` | Red fill, white font |
| Top 10% of 2021 unit sales | `L2:L41` | Blue fill |

The green CAGR condition uses a numeric check so text values such as `Formula Error` are not mistakenly classified as high-growth percentages. See the [findings document](docs/01-conditional-formatting-findings.md) for the exact account-level details and formula logic.

</details>


## Completed exercise: Formulas and Functions

<details open>
<summary><strong>Task 02 — Put Formulas Into Practice</strong> · <code>Completed</code></summary>

This exercise applied six account-level calculations and two summary functions to the **Excel Formula Practice** workbook. The completed workbook is stored under `material/02-formulas-and-functions/`, with the formula logic and learning outcomes documented below.

**Workbooks:**
- [Original practice workbook](material/02-formulas-and-functions/Excel_Formula_Practice.xlsx)
- [Completed workbook](material/02-formulas-and-functions/Excel_Formula_Practice_Completed.xlsx)

**Findings:** [Task 02 — Formulas and Functions Findings](docs/02-formulas-and-functions-findings.md)

### Account-level calculations

The formulas below use the worksheet layout documented during the exercise: annual unit sales from 2016–2021 in columns G:L, with the first account on row 2.

| Calculation | Formula | Purpose |
|---|---|---|
| Total Unit Sales | `=SUM(G2:L2)` | Adds annual unit sales from 2016 through 2021. |
| Average Annual Unit Sales | `=AVERAGE(G2:L2)` | Calculates average annual unit sales. |
| Highest Annual Unit Sales | `=MAX(G2:L2)` | Returns the highest annual unit-sales value. |
| Lowest Annual Unit Sales | `=MIN(G2:L2)` | Returns the lowest annual unit-sales value. |
| Sales Level | `=IF(L2>=$U$2,"High","Standard")` | Compares 2021 sales with the fixed threshold of 500 in U2. |
| Sales Change | `=IFERROR((L2-G2)/G2,"Check data")` | Calculates percentage change from 2016 to 2021 and handles errors. |

The Sales Level formula uses the absolute reference `$U$2` so the threshold does not shift when the formula is filled down. The Sales Change result is formatted as a percentage; a calculation error, including a zero 2016 baseline, returns `Check data`.

### Summary calculations

| Summary metric | Formula | Purpose |
|---|---|---|
| Accounts with Product 2 | `=COUNTIF(F2:F41,"Yes")` | Counts Product 2 cells containing `Yes`. |
| Retail 2021 Unit Sales | `=SUMIF(C2:C41,"Retail",L2:L41)` | Totals 2021 Unit Sales for Retail accounts. |

**Skills practiced:** `SUM`, `AVERAGE`, `MAX`, `MIN`, `IF`, absolute references, percentage change, `IFERROR`, `COUNTIF`, and `SUMIF`.

The completed calculations and documented findings have now been reviewed and verified by the learner in Excel. The findings document records the formulas, expected behavior, assumptions, and validation checks for future reference.

</details>


## Current exercise: Build Charts and Dashboards

<details open>
<summary><strong>Task 03 — Prepare Your Data for Visualization</strong> · <code>In progress</code></summary>

Task 03 Checkpoints 1–6 are complete based on the learner's reported Excel results. The `AccountSales` worksheet contains `A1:M41` (40 account records), the source data is an Excel Table named `AccountSales`, both required 2021 sales PivotTables and the 2016–2021 annual sales summary have been prepared, and all three required charts have been created: **Unit Sales by Year** (line), **Top 10 Accounts by 2021 Unit Sales** (bar), and **2021 Unit Sales by Account Type** (column). The updated working workbook is saved in the repository. Dashboard assembly and final review remain.

**Updated working workbook:** [Account Sales Data for Dashboard Practice — Working](material/03-build-charts-and-dashboards/Account_Sales_Data_for_Dashboard_Practice_Working.xlsx) · [Original practice workbook](material/03-build-charts-and-dashboards/Account_Sales_Data_for_Dashboard_Practice.xlsx)

**Preparation guide:** [Task 03 — Charts and Dashboard Preparation](docs/03-chart-and-dashboard-preparation.md)

### Current objectives

1. Save a separate working copy of the supplied workbook and preserve the original.
2. <img src="https://img.shields.io/badge/Complete-2E8B57?style=flat-square" alt="Complete" /> Inspect the `AccountSales` worksheet: range `A1:M41`, 40 account records.
3. <img src="https://img.shields.io/badge/Complete-2E8B57?style=flat-square" alt="Complete" /> Create the Excel Table named `AccountSales`.
4. <img src="https://img.shields.io/badge/Complete-2E8B57?style=flat-square" alt="Complete" /> Create and validate PivotTables for 2021 sales by account type and the top 10 accounts.
5. <img src="https://img.shields.io/badge/Complete-2E8B57?style=flat-square" alt="Complete" /> Prepare the 2016–2021 annual sales summary and create the **Unit Sales by Year** line chart.
6. <img src="https://img.shields.io/badge/Complete-2E8B57?style=flat-square" alt="Complete" /> Create and format all three required charts:
   - Unit Sales by Year — line chart.
   - Top 10 Accounts by 2021 Unit Sales — bar chart.
   - 2021 Unit Sales by Account Type — column chart.
7. Build the `Dashboard` worksheet, arrange all three charts, and complete the final review before submission.

Checkpoints 1–6 are complete based on learner-reported Excel results. The PivotTable summaries add to their displayed totals: 17,278 units by account type and 8,221 units for the Top 10 accounts. The annual summary reports growth from 9,773 units in 2016 to 17,278 in 2021 (+76.8%). The learner confirms that all three charts are complete: the annual trend line chart, the Top 10 account bar chart (Cloudline Commerce at the top), and the account-type column chart (Online highest). The updated working workbook is saved in the repository. The Dashboard worksheet, final validation, and submission remain outstanding.

</details>

## Repository guide

The repository separates learning documentation from the Excel files used during practice.

~~~text
forage-excel-skills-job-simulation/
│
├── README.md
├── LEARNING_LOG.md
├── .gitignore
│
├── docs/
│   ├── Excel Keyboard Shortcuts.docx
│   ├── Excel Formula Quick Reference Guide.docx
│   ├── 01-conditional-formatting-findings.md
│   ├── 02-formulas-and-functions-findings.md
│   └── 03-chart-and-dashboard-preparation.md
│
├── learning-notes/
│   ├── data-preparation.md
│   ├── formulas-and-analysis.md
│   ├── charts-and-dashboards.md
│   ├── automation-and-vba.md
│   └── reflection.md
│
└── material/
    ├── 01-conditional-formatting/
    │   └── Account Sales Data for Analysis.xlsx
    ├── 02-formulas-and-functions/
    │   ├── Excel_Formula_Practice.xlsx
    │   └── Excel_Formula_Practice_Completed.xlsx
    └── 03-build-charts-and-dashboards/
        ├── Account_Sales_Data_for_Dashboard_Practice.xlsx
        └── Account_Sales_Data_for_Dashboard_Practice_Working.xlsx
~~~

### Quick navigation

| Resource | Purpose |
|---|---|
| [Learning log](LEARNING_LOG.md) | Record study sessions, takeaways, and next steps. |
| [Excel keyboard shortcuts](docs/Excel%20Keyboard%20Shortcuts.docx) | Quick reference for navigating and working efficiently in Excel. |
| [Excel formula quick reference](docs/Excel%20Formula%20Quick%20Reference%20Guide.docx) | Formula syntax and examples to support spreadsheet analysis. |
| [Data preparation](learning-notes/data-preparation.md) | Document workbook exploration, data-quality findings, and cleanup decisions. |
| [Formulas and analysis](learning-notes/formulas-and-analysis.md) | Record formulas, assumptions, and validation checks. |
| [Charts and dashboards](learning-notes/charts-and-dashboards.md) | Capture KPI definitions, visualization choices, and insights. |
| [Automation and VBA](learning-notes/automation-and-vba.md) | Document original automation logic and tests. |
| [Reflection](learning-notes/reflection.md) | Summarize learning outcomes and improvement opportunities. |
| [Task 01 findings](docs/01-conditional-formatting-findings.md) | Review the dataset findings, exact affected accounts, conditional-formatting logic, and key takeaways. |
| [Task 02 findings](docs/02-formulas-and-functions-findings.md) | Review the account-level formulas, summary calculations, assumptions, and validation checklist. |
| [Task 03 preparation guide](docs/03-chart-and-dashboard-preparation.md) | Follow the Excel Table, PivotTable, chart-selection, and dashboard workflow. |
| [Task 03 updated working workbook](material/03-build-charts-and-dashboards/Account_Sales_Data_for_Dashboard_Practice_Working.xlsx) | Open the learner's updated workbook containing the three completed charts. |
| [Task 03 original practice workbook](material/03-build-charts-and-dashboards/Account_Sales_Data_for_Dashboard_Practice.xlsx) | Open the original account sales practice workbook. |
| [Completed Task 02 workbook](material/02-formulas-and-functions/Excel_Formula_Practice_Completed.xlsx) | Open the completed formula-practice workbook. |
| [Completed Task 01 workbook](material/01-conditional-formatting/Account_Sales_Data_Analysis_Completed.xlsx.xlsx) | Open the completed conditional-formatting workbook. |
| [Practice materials](material/) | Browse the Excel workbooks used for the exercises. |

## How I work through each exercise

1. **Inspect first.** Understand the workbook, data grain, columns, data types, and task objective before editing.
2. **Use Excel efficiently.** Prefer keyboard shortcuts and built-in Excel GUI features where they are appropriate.
3. **Apply the technique.** Use the required rules, formulas, charts, or automation without changing source data unintentionally.
4. **Validate the result.** Check ranges, formulas, edge cases, totals, and formatting behavior.
5. **Document the reasoning.** Record what was done, why it was done, what was observed, and what remains uncertain.
6. **Publish responsibly.** Add original work and permitted workbook outputs, with sensitive information removed.

## Skills in scope

<p>
  <img src="https://img.shields.io/badge/Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white" alt="Excel" />
  <img src="https://img.shields.io/badge/Data%20Quality-4472C4?style=flat-square" alt="Data quality" />
  <img src="https://img.shields.io/badge/Conditional%20Formatting-8A63D2?style=flat-square" alt="Conditional formatting" />
  <img src="https://img.shields.io/badge/Excel%20Formulas-217346?style=flat-square" alt="Excel formulas" />
  <img src="https://img.shields.io/badge/Data%20Analysis-0F766E?style=flat-square" alt="Data analysis" />
  <img src="https://img.shields.io/badge/Data%20Visualization-2563EB?style=flat-square" alt="Data visualization" />
  <img src="https://img.shields.io/badge/Dashboard%20Design-475569?style=flat-square" alt="Dashboard design" />
  <img src="https://img.shields.io/badge/VBA%20%26%20Macros-795548?style=flat-square" alt="VBA and macros" />
  <img src="https://img.shields.io/badge/Business%20Communication-B45309?style=flat-square" alt="Business communication" />
</p>

These are the skills covered by the learning plan; the presence of a badge does not mean every topic has already been completed.

## Evidence and sharing principles

- Keep the original practice workbook unchanged and use a separate working copy when completing an exercise.
- Add only original analysis, notes, formulas, screenshots, or macros that I created.
- Check the provider's terms before publicly redistributing supplied datasets, workbooks, templates, or other course materials.
- Remove personal information, account details, credentials, and certificate verification codes before publishing.
- Explain the purpose and behavior of any macro; do not ask reviewers to enable macros without context.
- Mark an exercise complete only after its output has been reviewed and validated.

## Progress

The overall simulation is **in progress**. Task 01 — Conditional Formatting and Task 02 — Put Formulas Into Practice are completed and documented. For Task 03, the data preparation, PivotTables, annual summary, and all three required charts are complete; dashboard assembly and final review remain.

For session-by-session updates, see [LEARNING_LOG.md](LEARNING_LOG.md).

---

<div align="center">

**Learning by doing. Validating every result. Documenting the reasoning.**

[Forage Academy](https://www.theforage.com/) · [Back to repository](https://github.com/Sohila-Khaled-Abbas/forage-excel-skills-job-simulation)

</div>
