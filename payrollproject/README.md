# Burrows Island Light Station Multi-Source Payroll Analytics Dashboard

**ETL → KPI Dashboard**: Aggregates contractor, expense, and input data into automated performance metrics, cost allocation, and compliance reporting for Burrows Island Light Station restoration project.

## Dashboard Preview

<img width="768" height="402" alt="image" src="https://github.com/user-attachments/assets/9811b6a3-08ca-4970-a402-d5ea1fa9fa3d" />

**Key Metrics Delivered:**  

    **Total Hours** | **Total Payroll** | **Average Rate** | **Total Expenses** | **Contractor Hours** | **Contractor Paid** | **Avg Hours/Job**

## Skills Demonstrated

- **ETL Pipeline**: Multi-sheet integration (Contractors + Expenses + Data Input)
- **KPI Dashboards**: SUMIF/SUMPRODUCT for contractor rankings and totals
- **Data Validation**: Dropdown lists eliminate entry errors
- **Dynamic Reporting**: UNIQUE arrays, VLOOKUP automation, real-time updates

## Business Case

**Raw Data**: Burrows Island restoration project payroll and expenses across multiple contractors and phases  
**Objective**: Build automated KPI dashboard for grant compliance and cost visibility  
**Solution**: 4 integrated sheets → real-time contractor performance and phase analysis  
**Outcome**: Audit-ready reporting with zero manual calculation errors

## Table of Contents  
#### Project Overview
- [Dashboard Preview](#dashboard-preview)
- [Skills Demonstrated](#skills-demonstrated)
- [Business Case](#business-case)  

#### System Setup and Instructions
- [Contractors Sheet Setup](#1-contractors-sheet-setup)  
- [Expenses Sheet Setup](#2-expenses-sheet-setup)  
- [Data Input Sheet Setup](#3-data-input-sheet-setup) 
- [KPI Dashboard](#4-kpi-dashboard)
- [Key Formulas – Copy & Paste](#5-key-formulas--copy--paste-ready)  

#### Business Intelligence & Analytics
- [Expanded Payroll Framework for Restoration Projects](#expanded-payroll-framework-for-restoration-projects)
- [BI Skills Demonstrated](#bi-skills-demonstrated)

***

## 1. Contractors Sheet Setup  
**Purpose**: Master reference for lookup and validation

### Step 1: Create the Sheet  
Right‑click sheet tab → **Rename → Contractors**  
Merge and center `A1:D1`  
`A1:` **Burrows Island Light Stand Payroll Summary** (bold, 12 pt)  
**View → Freeze Panes → Top Row**

### Step 2: Headers  
| Contractor Name | Specialty | Hourly Rate | Contact Email | Total Hours |
|------------------|------------|--------------|----------------|-------------|

### Step 3: Example Data  
| Contractor Name | Specialty | Hourly Rate | Contact | Hours |
|------------------|------------|-------------|----------|--------|
| Barnet Roofing | Roofer | $50.00 | contact@roof.com | 180 |
| Coast Guard Cleanup | Cleanup Specialist | $60.00 | contracts@uscg.mil | 240 |
| Salish Sea Volunteers | Trailer Worker | $0.00 | volunteers@salishsea.org | 320 |
| NWSS Preservation | Historic Preservation Specialist | $75.00 | nwss@burrowslighthouse.com | 150 |

***

## 2. Expenses Sheet Setup  
**Granular cost tracking by phase and material**

**Sheet Name:** `Expenses`

| Date | Phase | Site Area | Material Type | Contractor | Hours | Cost | Notes |
|------|-------|------------|----------------|-------------|--------|------|--------|
| 1/15/2026 | Preservation | Keepers Quarters | Timber | Barnet Roofing | 40 | $2,000 | Roof timber replacement |
| 1/20/2026 | Cleanup | Lighthouse Tower | Paint Removal | Coast Guard Cleanup | 32 | $1,920 | Lead paint removal |
| 1/25/2025 | Assessment | Boathouse | Stone | Heritage Historians | 16 | $1,440 | 1906 foundation survey |

**Formatting:**  
- Column A → Short Date  
- Column G → Currency  

***

## 3. Data Input Sheet Setup
**Primary transactional data with validation**

**Sheet Name:** `Data Input`

| Date | Contractor | Hours Worked | Hourly Rate | Job Type | Notes |
|-------|-------------|---------------|--------------|----------|--------|
| 2026‑02‑05 | NWSS Preservation | 24.0 | $75.00 | Restoration | — |
| 2026‑02‑10 | Island Marine Service | 12.0 | $80.00 | Logistics | — |

### Dropdown Lists  
**Contractor (Column B)** → Data Validation → List  
Source: `=INDIRECT("Contractors!A3:A10")`  

**Job Type (Column E)** → Data Validation → List  
Source: `Restoration,Welding,Painting,Electrical,Carpentry,Masonry`

***

## 4. KPI Dashboard 
**Summary Sheet** - Real-time aggregation across all data sources:

**Sheet Name:** `Summary`

| Metric | Formula |
|---------|----------|
| Total Hours | `=SUM('Data Input'!C:C)` |
| Total Payroll | `=SUMPRODUCT('Data Input'!C:C,'Data Input'!D:D)` |
| Average Rate | `=AVERAGE('Data Input'!D:D)` |
| Total Expenses | `=SUM('Expenses'!G:G)` |

**Contractor Summary Table**
| Contractor | Hours | Paid | Avg Hours/Job |
|-------------|--------|------|----------------|
| `=UNIQUE('Data Input'!B:B)` | `=SUMIF('Data Input'!B:B,$A9,'Data Input'!C:C)` | `=SUMPRODUCT(('Data Input'!B:B=$A9)*('Data Input'!C:C)*('Data Input'!D:D))` | `=$B9/COUNTIF('Data Input'!B:B,$A9)` |

***

## 5. Key Formulas – Copy & Paste Ready  
```excel
'Payroll Totals
=SUM('Data Input'!C:C)
=SUMPRODUCT('Data Input'!C:C,'Data Input'!D:D)

'Per Contractor
=SUMIF('Data Input'!B:B,$A9,'Data Input'!C:C)
=SUMPRODUCT(('Data Input'!B:B=$A9)*('Data Input'!C:C)*('Data Input'!D:D))

'Auto Rate Lookup
=VLOOKUP(B2,Contractors!A:C,3,FALSE)

'Total Expenses
=SUM('Expenses'!G:G)
```
## Expanded Payroll Framework for Restoration Projects

**6 strategic enhancements that transform basic payroll into grant-compliant project management.**

### Recommended Enhancements

| Feature | Purpose | Formula / Example |
|----------|----------|-------------------|
| **Grant Funding Source** | Track which grant (e.g., *Lighthouse Environmental Program* or preservation fund) covers each contractor’s pay — critical for audits and reimbursements. | `=SUMIF(GrantColumn,"LEP",PayrollRange)` |
| **Materials & Expense Allocation** | Link payroll directly to project costs like materials, permits, transportation, or dumping fees to reveal full restoration cost per site area. | `=SUMIF(JobTypeRange,"Roofing",ExpenseRange)` |
| **Volunteer Hours Offset** | Log volunteer contributions separate from paid labor to properly record *in‑kind* support required by certain grants. | Use separate sheet or “Volunteer” tag in column |
| **Prevailing Wage Certification** | Verify that contractors meet federal/state Davis‑Bacon wage rates for skilled trades in public projects. | `=VLOOKUP(Trade,WageTable,2,FALSE)` |
| **Safety & Training Logs** | Record certifications like OSHA or site‑specific hazard training (e.g., contaminated soil handling). | Checkbox/Dropdown field |
| **Permit References** | Cross‑reference hours and activities with corresponding restoration or environmental permits. | Text reference or hyperlink field |

### Why These Matter for Restoration Projects
These project enhancements would make the payroll tracker especially useful for **historic preservation and grant‑funded restoration projects**, such as lighthouse or heritage site work.

**Standard payroll** (name, hours, rate, total) works for simple jobs. **Restoration projects** need robust documentation for:

Standard payroll tables (contractor, hours, rate, total) fall short for restoration work, which demands **robust documentation** for grant oversight, audits, and preservation compliance. **Grant Funding Source** columns link labor to specific programs (state awards, Lighthouse Environmental Program) for accountability. **Materials & Expense Allocation** connects payroll to permits, ferry costs, and period materials for true cost visibility. **Volunteer Hours** separate in-kind contributions essential for grant documentation. **Compliance & Safety Fields** (prevailing wage, OSHA, permits) ensure state/federal requirements are met—transforming basic Excel into a **comprehensive financial/compliance tool** for heritage restoration complexity.

**Result:** Basic Excel sheet → **Comprehensive financial/compliance tool** built for heritage restoration complexity.

### BI Skills Demonstrated

**Skills beyond formulas:**
- Multi-source data aggregation and analysis
- Creating audit-ready dashboards and reports
- Cost allocation and KPI tracking across categories
- Data validation and stakeholder-ready visualizations

**Perfect for:** Data Analysts, BI Analysts, Reporting Specialists, Analytics roles
