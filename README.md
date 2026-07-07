# 📊 Employee Turnover Analysis Dashboard
### Excel-Based HR Analytics | Personal Portfolio Project

> An end-to-end HR analytics project built in Microsoft Excel, covering synthetic data generation, data cleaning, dashboard design, pivot-based analysis, and VBA automation — simulating a real-world People Analytics workflow for a medium-sized Philippine-based company.

---

## 📌 Table of Contents

- [Project Overview](#project-overview)
- [Objectives](#objectives)
- [Tools & Technologies](#tools--technologies)
- [Dataset](#dataset)
- [Project Architecture](#project-architecture)
- [Dashboard Features](#dashboard-features)
- [VBA Automation](#vba-automation)
- [Key Findings](#key-findings)
- [Data Cleaning Process](#data-cleaning-process)
- [Limitations](#limitations)
- [Skills Demonstrated](#skills-demonstrated)
- [How to Use](#how-to-use)
- [Author](#author)

---

## Project Overview

This project simulates an Employee Turnover Analysis that an HR Data Analyst might produce for a medium-sized company. Starting from a synthetic but realistic dataset of 1,000+ employee records spanning 2018–2026, the project covers the full analyst workflow: raw data → cleaning → analysis → visualization → automation → reporting.

The dashboard tracks employee headcount, turnover rates, exit patterns, tenure distributions, and department-level attrition — with interactive slicers, auto-updating pivot charts, a live mirrored data table, and VBA-powered record management forms.

---

## Objectives

- Analyze employee turnover trends across departments, time periods, gender, and age brackets
- Identify patterns in exit reasons and tenure before exit
- Build an interactive, scalable dashboard usable by non-technical HR staff
- Automate record management (add, update, delete) via VBA UserForms
- Demonstrate end-to-end data analyst skills suitable for a Junior DA role

---

## Tools & Technologies

| Tool | Purpose |
|---|---|
| **Microsoft Excel** | Primary tool — data storage, analysis, dashboard, automation |
| **Power Query** | Data import, type transformation, date parsing |
| **PivotTables & PivotCharts** | Aggregations and interactive visualizations |
| **Excel Formulas** | `COUNTIFS`, `INDEX`, `IFERROR`, `SUMPRODUCT`, dynamic arrays |
| **VBA / Macros** | UserForms, record management, slicer-mirror sync |
| **Synthetic Dataset** | AI-generated (Claude/Anthropic) to mirror real HR data characteristics, including intentional data quality issues for cleaning practice |
| **PowerPoint** | UserForm background design assets |
| **GitHub** | Version control and portfolio hosting |

---

## Dataset

### Source
A synthetic dataset was generated using AI assistance (Claude/Anthropic) to simulate realistic HR records for a Philippine-based company. The dataset was designed to mirror real-world HR data characteristics — including realistic Filipino names, proportional department distributions, date ranges, and intentional data quality issues — providing a practical foundation for authentic data cleaning and analysis work.

### Specifications

| Attribute | Detail |
|---|---|
| Total Records | ~1,135 rows (including intentional duplicates) |
| Clean Records | ~1,081 unique employees after data cleaning |
| Date Range | January 2018 – June 2026 |
| Departments | Customer Support, IT, Marketing, HR, Sales, Finance |
| Status | Active (667) / Exited (414) |

### Columns

| Column | Description |
|---|---|
| `Employee ID` | Unique identifier (auto-incremented from 1001) |
| `Full Name` | Realistic Filipino names (First Middle Last format) |
| `Department` | One of six departments |
| `Gender` | Male / Female / Other |
| `Age` | Integer, range 21–65 |
| `Age Bracket` | Derived: 20s / 30s / 40s / 50s / 60s |
| `Join Date` | Date employee was hired (2018–2025) |
| `Exit Date` | Date employee left (blank if still Active) |
| `Exit Reason` | One of eight categories (blank if Active) |
| `Status` | Active or Exited |
| `Tenure (days)` | Calculated: Exit Date − Join Date (or today if Active) |
| `Tenure (years)` | Derived integer from Tenure (days) |
| `Tenure (yrs. mos. days)` | Human-readable tenure string |

### Intentional Data Quality Issues (for cleaning practice)
- **40 exact duplicate rows** — identical across all fields
- **15 near-duplicate rows** — same Employee ID but one altered field (Department, Gender, or Age) to simulate real-world data entry errors

---

## Project Architecture

```
employee-tracking.xlsm
│
├── DASHBOARD          ← Interactive dashboard + live mirror data table
├── PIVOT TABLES       ← All PivotTables and helper calculations (hidden)
├── Employee Data      ← Primary data table (EmployeeData)
└── helpers            ← KPI source cells, turnover rate helper table
```

### Data Flow

```
Python CSV generation
        ↓
Power Query import + date type fix
        ↓
Employee Data sheet (EmployeeData Table)
        ↓
    ┌───────────────────────────────┐
    │                               │
PivotTables                  INDEX/IFERROR
(PIVOT TABLES sheet)         Mirror Table
    │                               │
PivotCharts               Live mirror on Dashboard
(Dashboard)               synced to Department Slicer
    │                               │
    └──────── Department Slicer ────┘
              (VBA bridge macro)
```

---

## Dashboard Features

### KPI Cards
- **Total Headcount** — total unique employees across all years
- **Turnover Rate** — overall rate (Total Exited ÷ Total Headcount)
- **Active Employees** — current active headcount
- **Exited Employees** — total exits

### Charts & Visuals

| Visual | Type | Insight |
|---|---|---|
| Employee Exits per Department | Horizontal Bar | Which departments have highest attrition |
| Exit Reason Breakdown | Bar Chart | Distribution of why employees leave |
| New Hires vs. Exits Over Time | Dual Line | Monthly hiring and attrition trend |
| Company Turnover Rate Over Time | Line | Quarterly turnover rate % trend (2018–Q2 2026) |
| Average Exited Tenure | Histogram | Tenure distribution of employees who left |
| Average Active Tenure | Histogram | Tenure distribution of current employees |
| Active & Exited by Age Bracket | Grouped Bar | Attrition pattern across age groups |
| Employee Headcount YoY | Table | Year-over-year headcount and turnover summary |

### Interactive Elements
- **Department Slicer** — filters all PivotCharts and mirror data table simultaneously
- **Live Mirror Data Table** — reflects full `EmployeeData` table using `INDEX/IFERROR` formulas, styled to match dashboard theme, synced to Department Slicer via VBA

---

## VBA Automation

Three record management features are accessible via buttons on the dashboard, built as VBA UserForms with input validation, confirmation dialogs, and safeguards against data loss.

### Add New Employee (`frmAddEmployee`)
- Auto-generates the next Employee ID
- Dropdown validation for Department and Gender
- Date format parsing (`MM/DD/YYYY`) using a shared `ParseDate_MMDDYYYY()` utility
- Validates age range (18–70) and join date not in the future
- Writes directly to `EmployeeData` table with Status hardcoded as "Active"

### Update Record (`frmUpdateRecord`)
- Two-stage employee lookup: type-to-filter search TextBox + filtered ComboBox results
- Displays current employee details for review before editing
- Conditionally reveals Exit Date and Exit Reason fields only when Status is changed to "Exited"
- Validates exit date is after join date and not in the future
- Confirmation dialog shows full change summary before writing
- Handles reverting an Exited employee back to Active (clears exit fields)

### Delete Record
- Scoped as a data correction tool (duplicate/erroneous entry removal only)
- Three-step confirmation: review full record → retype Employee ID → final irreversible confirm
- Uses `vbCritical` warning to signal permanent action

### Refresh Dashboard
- One-click button refreshes all PivotTables via `ThisWorkbook.RefreshAll`

### Slicer–Mirror Table Bridge
- `Worksheet_PivotTableUpdate` event on the PIVOT TABLES sheet detects slicer changes
- `SyncMirrorTableFilter` reads selected items from `SlicerCache` and applies `AutoFilter` to the mirror table
- Supports multi-select slicer filtering using `xlFilterValues`

### Shared Module Utilities (`Module1`)
| Function | Purpose |
|---|---|
| `FindEmployeeRow()` | Reusable lookup: finds a table row by Employee ID |
| `ParseDate_MMDDYYYY()` | Explicit date parsing bypassing regional settings ambiguity |
| `SyncMirrorTableFilter()` | Bridges Department Slicer to mirror table AutoFilter |

---

## Key Findings

> Note: These findings are based on synthetic data and are illustrative of the analytical approach rather than real organizational insights.

- **Overall turnover rate: 38.3%** across the full 2018–2026 period
- **Customer Support and IT** have the highest absolute exit counts (100 and 96 respectively), consistent with their larger department sizes
- **Better Opportunity** is the leading exit reason, suggesting competitive retention challenges
- **Average tenure before exit: 2 years 2 months** — exits occur early in the employee lifecycle, pointing to onboarding and early-tenure retention as priority areas
- **Average active tenure: 4 years 6 months** — employees who stay past the 2-year mark tend to remain significantly longer
- **Turnover rate increased steadily from 2% in 2018 to a peak of 16.1% in 2025**, reflecting either organizational growth pains or worsening retention — worth investigating alongside external factors
- **50s age bracket** has the highest exit count among age groups, driven partly by Retirement as an exit reason
- **2025 saw the highest single-year exits (116)** — a significant spike worth further investigation in a real dataset

---

## Data Cleaning Process

| Step | Issue | Action Taken |
|---|---|---|
| 1 | Date columns imported as text (`DD/MM/YYYY`) | Fixed via Power Query: Change Type > Using Locale (English UK) |
| 2 | 40 exact duplicate rows | Identified via `COUNTIF` on Employee ID, removed using Power Query |
| 3 | 15 near-duplicate rows (altered field) | Identified by sorting on Employee ID and visually comparing flagged pairs, removed after verification |
| 4 | Exit Date/Exit Reason showing `0` in mirror table for Active employees | Fixed via `IF(...=0, "", ...)` wrapper in `INDEX/IFERROR` mirror formulas |
| 5 | Date serial numbers displaying in mirror table | Fixed via `Format Cells > Custom > MM/DD/YYYY` on full 5,000-row buffer |
| 6 | Power Query refresh risk overwriting manual edits | Resolved by severing Power Query connection after cleaning (Keep Data, Delete Query) |

---

## Limitations

- **Synthetic dataset** — findings reflect generated patterns, not real organizational behavior. Some distributions (e.g., exit reasons) are roughly equal by design rather than reflecting natural skew
- **2026 is a partial year** — Q1–Q2 2026 data exists for exits but not new hires, since the generation cutoff was mid-2026. Charts and turnover rates for 2026 should be interpreted as partial-period figures
- **Static after severance** — once the Power Query connection was severed to enable manual record management, the dataset no longer auto-refreshes from an external source. In a real deployment, this would be replaced by a proper database connection or SharePoint-linked data source
- **VBA macro security** — the `.xlsm` file requires macros to be enabled. In a real corporate environment, this would go through IT security review and code signing
- **Turnover Rate helper table** — the quarterly turnover rate chart is built on a manually maintained helper table rather than a fully dynamic PivotTable, since point-in-time headcount snapshots aren't natively supported by PivotTable aggregations

---

## Skills Demonstrated

**Data Analysis**
- Exploratory data analysis and pattern identification
- HR metrics: turnover rate, tenure analysis, headcount tracking, cohort comparison
- Time-series analysis (quarterly trend, YoY comparison)
- Cross-dimensional analysis (department × exit reason, age × status)

**Excel**
- PivotTables, PivotCharts, Slicers
- Dynamic array formulas (`INDEX`, `IFERROR`, `COUNTIFS`, `SUMPRODUCT`)
- Power Query (import, type transformation, locale-aware date parsing)
- Named Tables and structured references
- Custom date formatting and number formats
- Conditional formatting

**VBA / Programming**
- UserForm design and event-driven programming
- Input validation and error handling
- `ListObject` / `ListRow` table manipulation
- `SlicerCache` reading and `AutoFilter` programmatic control
- Module organization and shared utility functions
- Worksheet event procedures

**Data Storytelling & Design**
- Dashboard layout and visual hierarchy
- Chart selection rationale (why each chart type fits its metric)
- Slicer-driven interactivity for non-technical users
- README and project documentation

---

## How to Use

### Requirements
- Microsoft Excel 2016 or later (Microsoft 365 recommended)
- Macros must be enabled (File > Options > Trust Center > Macro Settings > Enable all macros)

### Getting Started
1. Download `employee-tracking.xlsm` from this repository
2. Open in Excel and click **Enable Macros** when prompted
3. Navigate to the **DASHBOARD** sheet
4. Use the **Department Slicer** to filter visuals and the data table by department
5. Use the dashboard buttons to **Add Employee**, **Update Record**, **Delete Record**, or **Refresh Dashboard**

### Adding a New Employee
1. Click **Add Employee** on the dashboard
2. Fill in the form fields (Employee ID is auto-generated)
3. Click **Add Employee** to confirm — the new record appears in the data table and mirror immediately

### Updating an Employee Record (e.g., marking as Exited)
1. Click **Update Record** on the dashboard
2. Type the employee's name or ID in the search box
3. Select the employee from the filtered dropdown
4. Change the Status to **Exited** — Exit Date and Exit Reason fields will appear
5. Fill in the exit details and click **Update Record**

### Deleting a Record (data correction only)
1. Click **Delete Record** on the dashboard
2. Enter the Employee ID to delete
3. Review the full record details shown
4. Retype the Employee ID to confirm
5. Confirm the final irreversible deletion prompt

---

## Author

**J.M.L. Villanueva**
📧 jmateo.vill@gmail.com
🎓 Bachelor of Science in Computer Science
📍 Philippines

*This project was built as a personal portfolio piece to demonstrate data analysis, Excel, and VBA skills aligned with Junior Data Analyst role requirements.*

---

> *"Data is only as useful as the decisions it enables."*
