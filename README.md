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
- [Current exercise: Conditional Formatting](#current-exercise-conditional-formatting)
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

## Current exercise: Conditional Formatting

<details open>
<summary><strong>Task 01 — Put Conditional Formatting Into Practice</strong> · <code>In progress</code></summary>

The first exercise uses the **Account Sales Data for Analysis** workbook to identify data-quality issues and highlight commercially important values through conditional formatting.

**Workbook:** [<code>Account Sales Data for Analysis.xlsx</code>](material/01-conditional-formatting/Account%20Sales%20Data%20for%20Analysis.xlsx)

### Required formatting rules

| # | What to identify | Required formatting |
|---:|---|---|
| 1 | Formula errors | Purple fill with white text |
| 2 | Missing values | Yellow fill |
| 3 | Accounts not cross-sold with Product 2 | Orange fill on the Product 2 cells |
| 4 | 5-year CAGR values of 100% or higher | Green fill |
| 5 | Negative 5-year CAGR values | Red fill with white text |
| 6 | Top 10% of 2021 unit sales | Blue fill |

### What this exercise is intended to demonstrate

- Inspect a workbook before changing its formatting.
- Distinguish a genuinely missing value from a valid zero or a business status such as “not purchased”.
- Apply built-in conditional formatting and formula-based rules where needed.
- Scope each rule to the correct column or data range.
- Review the rules manager and validate the final workbook without changing source values.

**Progress note:** The source practice workbook is stored in <code>material/</code>. The formatting work should only be marked complete after the rules have been applied and checked in the working workbook.

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
│   └── Excel Keyboard Shortcuts.docx
│
├── learning-notes/
│   ├── data-preparation.md
│   ├── formulas-and-analysis.md
│   ├── charts-and-dashboards.md
│   ├── automation-and-vba.md
│   └── reflection.md
│
└── material/
    └── 01-conditional-formatting/
        └── Account Sales Data for Analysis.xlsx
~~~

### Quick navigation

| Resource | Purpose |
|---|---|
| [Learning log](LEARNING_LOG.md) | Record study sessions, takeaways, and next steps. |
| [Excel keyboard shortcuts](docs/Excel%20Keyboard%20Shortcuts.docx) | Quick reference for navigating and working efficiently in Excel. |
| [Data preparation](learning-notes/data-preparation.md) | Document workbook exploration, data-quality findings, and cleanup decisions. |
| [Formulas and analysis](learning-notes/formulas-and-analysis.md) | Record formulas, assumptions, and validation checks. |
| [Charts and dashboards](learning-notes/charts-and-dashboards.md) | Capture KPI definitions, visualization choices, and insights. |
| [Automation and VBA](learning-notes/automation-and-vba.md) | Document original automation logic and tests. |
| [Reflection](learning-notes/reflection.md) | Summarize learning outcomes and improvement opportunities. |
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

The project is **in progress**. The first exercise's source workbook and the Excel shortcut reference are available in the repository. The task notes and completed workbook outputs will be updated as the work is performed and verified.

For session-by-session updates, see [LEARNING_LOG.md](LEARNING_LOG.md).

---

<div align="center">

**Learning by doing. Validating every result. Documenting the reasoning.**

[Forage Academy](https://www.theforage.com/) · [Back to repository](https://github.com/Sohila-Khaled-Abbas/forage-excel-skills-job-simulation)

</div>
