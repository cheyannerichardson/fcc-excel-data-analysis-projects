# Excel Payroll Tracker – Complete Step-by-Step Guide
### Burrows Island Light Stand Restoration Project

A fully functional Excel-based payroll and expense tracking system inspired by FreeCodeCamp’s beginner-friendly tutorials. This guide walks you through worksheet setup -- from contractor data to automate summary formulas.

## 📘 Project Overview

This project demonstrates how to:
  - Organize payroll data for multiple contractors
  - Use Excel formulas for automation ('SUMIF', VLOOKUP', 'SUMPRODUCT')
  - Create validatinon lists for consistent data entry
  - Build a summary dashboard for total payroll hours and expenses

 
## 🗂️ Table of Contents  
1. [Excel Basics Every Beginner Needs](#1-excel-basics-every-beginner-needs)  
2. [Contractors Sheet Setup](#2-contractors-sheet-setup)  
3. [Data Input Sheet Setup](#3-data-input-sheet-setup)  
4. [Expenses Sheet Setup](#4-expenses-sheet-setup)  
5. [Summary Sheet Setup](#5-summary-sheet-setup)  
6. [Key Formulas – Copy & Paste](#6-key-formulas--copy--paste-ready)  
7. [Expanded Payroll Framework for Restoration Projects](#7-expanded-payroll-framework-for-restoration-projects)

***

## 1. Excel Basics Every Beginner Needs
- **Cell:** One box (A1, B2, C3)  
- **Column:** Vertical (A, B, C …)  
- **Row:** Horizontal (1, 2, 3 …)  
- **Sheet:** Each tab at the bottom (e.g., “Sheet1”)  
- **Formula:** Starts with `=` sign  
- **Range:** Group of cells (e.g., `B:B`)

***

## 2. Contractors Sheet Setup  

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

## 3. Expenses Sheet Setup  

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

## 4. Data Input Sheet Setup  
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

## 5. Summary Sheet Setup  
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

## 6. Key Formulas – Copy & Paste Ready  
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

## 7. Expanded Payroll Framework for Restoration Projects

### 🔹 Recommended Enhancements

| Feature | Purpose | Formula / Example |
|----------|----------|-------------------|
| **Grant Funding Source** | Track which grant (e.g., *Lighthouse Environmental Program* or preservation fund) covers each contractor’s pay — critical for audits and reimbursements. | `=SUMIF(GrantColumn,"LEP",PayrollRange)` |
| **Materials & Expense Allocation** | Link payroll directly to project costs like materials, permits, transportation, or dumping fees to reveal full restoration cost per site area. | `=SUMIF(JobTypeRange,"Roofing",ExpenseRange)` |
| **Volunteer Hours Offset** | Log volunteer contributions separate from paid labor to properly record *in‑kind* support required by certain grants. | Use separate sheet or “Volunteer” tag in column |
| **Prevailing Wage Certification** | Verify that contractors meet federal/state Davis‑Bacon wage rates for skilled trades in public projects. | `=VLOOKUP(Trade,WageTable,2,FALSE)` |
| **Safety & Training Logs** | Record certifications like OSHA or site‑specific hazard training (e.g., contaminated soil handling). | Checkbox/Dropdown field |
| **Permit References** | Cross‑reference hours and activities with corresponding restoration or environmental permits. | Text reference or hyperlink field |

### 🧩 Why These Additions Matter
These project enhancements would make the payroll tracker especially useful for 
**historic preservation and grant‑funded restoration projects**, such as lighthouse or heritage site work.

A standard payroll table typically includes columns for contractors names, hours worked, hourly rates, and total pay.
However, restoration wrok often required more robust documenation to satisy grant oversight, audit requirments, and preservation compliance.
Incorporating a **Grant Funding Source** column allows each labor cost to be linked to a specific grant or funding program (for example, 
a state preservation aware or the Lighthouse Enviromental Program), improving accountability and reimbursment tracking. Adding 
**Materials and Expense Allocations** fields helps connect payroll entries with related projects costs-such as permits, ferry transporation, or period-specific
construction materials-offering a more accurate understanding of total restoration expenditures.
Including a **Volunteer Hours** column enables seperation of paid labor from volunteer or in-kind constribuations, which is essential for grant documentation.
Finally, extending the tetable with **Compliance and Safety Fields**, such as pevailiing wage certifications, OSHA training verification, and 
permit reference numbers, ensures that the restoration effort meets state and federal requirments.

Together these would expand a basic Excel payroll sheet into a comprehensive financial and ocmplaince management tool [purpseo-built for 
the complex conditions of hertijage restoration projects.