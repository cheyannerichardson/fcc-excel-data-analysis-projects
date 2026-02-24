# Excel Payroll Tracker – Complete Step-by-Step Guide
### Burrows Island Light Stand Restoration Project

The Burrows Island Light Station near Anacortes, Washington is currently being restored through the efforts of the Keepers of the Burrows Island Light Station, preservation agencies, and skilled contractors. This Excel-based payroll and expense tracker models how structured data management can support heritage restoration projects by tracking labor, materials, and costs across multiple project phases. Such tools simplify coordination, improve transparency for grant compliance, and strengthen accountability among stakeholders. All figures used are hypothetical and for educational purposes only.

## 📘 Project Overview

This project demonstrates how to:
  - Organize payroll data for multiple contractors
  - Use Excel formulas for automation ('SUMIF', VLOOKUP', 'SUMPRODUCT')
  - Create validatinon lists for consistent data entry
  - Build a summary dashboard for total payroll hours and expenses

 
## 🗂️ Table of Contents  
#### Project Overview
- [Project Context](#project-contex)  

#### System Setup and Instructions
- [Excel Basics Every Beginner Needs](#1-excel-basics-every-beginner-needs)  
- [Contractors Sheet Setup](#2-contractors-sheet-setup)  
- [Expenses Sheet Setup](#3-expenses-sheet-setup)  
- [Data Input Sheet Setup](#4-data-input-sheet-setup) 
- [Summary Sheet Setup](#5-summary-sheet-setup)  
- [Key Formulas – Copy & Paste](#6-key-formulas--copy--paste-ready)  

#### Application and Use Cases
- [Expanded Payroll Framework for Restoration Projects](#7-expanded-payroll-framework-for-restoration-projects)  
- [Applying the System to Historical Restoration](#applying-the-system-to-historical-restoration)  
- [How to Use This Workbook for Your Own Projects](#how-to-use-this-workbook-for-your-own-projects)  

## Project Contex 
The Burrows Island Light Station, located just west of Anacortes, Washington, is a historically significant beacon first illuminated in 1906 to aid naval and commercial vessels navigating Rosario Strait. Today, its conservation is being led by the Keepers of the Burrows Island Light Station in partnership with local preservation groups, the Washington State Department of Archaeology and Historic Preservation, and specialized restoration contractors. This initiative focuses on structural stabilization, weatherproofing, and faithfully restoring the lighthouse and its surrounding buildings to preserve maritime heritage for future generations.

In a project of this scope, a payroll and expense tracking system provides a vital layer of organization and accountability. Historical restoration often involves multiple contractors and specialists—such as carpenters, stonemasons, electricians, and heritage consultants—working across distinct project phases. By using an organized Excel-based payroll tracker, managers can efficiently log labor hours, categorize work types, and track materials for each site area or task. This enables accurate cost estimation, phase-based budgeting, and transparent recordkeeping that aligns with grant eligibility and reporting requirements for heritage funding.

Beyond payroll management, modern data tools like this one streamline communication among teams, simplify compliance documentation, and ensure spending aligns with preservation priorities. They transform what might otherwise be a complex paper-based workflow into a clear, actionable dataset—helping restoration organizations like those at Burrows Island Light demonstrate financial integrity, plan future maintenance more effectively, and protect their legacy through responsible use of resources.

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
a state preservation award or the Lighthouse Enviromental Program), improving accountability and reimbursment tracking. Adding 
**Materials and Expense Allocations** fields helps connect payroll entries with related projects costs-such as permits, ferry transporation, or period-specific
construction materials-offering a more accurate understanding of total restoration expenditures.
Including a **Volunteer Hours** column enables seperation of paid labor from volunteer or in-kind contribuations, which is essential for grant documentation.
Finally, extending the table with **Compliance and Safety Fields**, such as pevailiing wage certifications, OSHA training verification, and 
permit reference numbers, ensures that the restoration effort meets state and federal requirments.

Together these would expand a basic Excel payroll sheet into a comprehensive financial and compliance management tool built for the complex conditions of hertijage restoration projects.

## Applying the System to Historical Restoration
A structured Excel-based data management system like this one can provide vital oversight in complex preservation projects such as the Burrows Island Light Station restoration. In real-world applications, heritage projects often involve multiple funding sources, specialized contractors, and strict documentation needs—all of which exceed the capacity of a basic payroll table. While a standard payroll tracker records contractor names, hours worked, hourly rates, and total pay, historical restoration work demands far more robust recordkeeping.

These project enhancements would make the payroll tracker especially useful for historic preservation and grant‑funded restoration efforts, including lighthouse or heritage site work. When applied to a real-world heritage initiative like the Burrows Island Light Station restoration, these expanded payroll features transform a basic spreadsheet into a dynamic project management tool. Linking each entry to a Grant Funding Source allows restoration managers to track which expenses align with specific funding programs—crucial for projects supported by multiple preservation grants. The Materials and Expense Allocation fields provide clarity on how labor ties to tangible restoration tasks, such as specialized masonry work, site safety preparation, or the transport of historic materials to the island. Tracking Volunteer Hours separately acknowledges community contributions, which are often necessary to document in-kind support for nonprofit audits and grant matching requirements. Meanwhile, Compliance and Safety Fields ensure that all work meets federal and state preservation standards while maintaining a transparent audit trail.


- **Improved transparency:** Links labor and expenses directly to funding sources, simplifying reimbursement and reporting to grant agencies.  
- **Compliance readiness:** Tracks certifications, permits, and safety standards to ensure alignment with state and federal regulations.  
- **Accurate cost allocation:** Connects payroll entries to material and expense categories for a full picture of restoration spending.  
- **Grant eligibility support:** Provides clear documentation needed for audits, reimbursement claims, and ongoing grant funding.  
- **Volunteer recognition:** Separates in‑kind and volunteer contributions, helping organizations quantify community support.  
- **Streamlined communication:** Consolidates data for preservation teams, funders, and contractors into one accessible system.  

Together, these additions expand a simple Excel payroll sheet into a comprehensive financial, compliance, and project‑tracking tool, purpose‑built for the complex conditions of heritage restoration. For the Burrows Island Light project, data management tools like this not only streamline administrative work but also strengthen stewardship of public funds, improve accountability across contractor teams, and preserve the integrity of one of Washington’s maritime landmarks. 

## How to Use This Workbook for Your Own Projects

A data management system like the **Burrows Island Light Station Payroll Tracker** demonstrates how spreadsheet tools can organize, record, and analyze complex restoration workflows.

**Contractors Sheet**: Teams can document each contractor’s trade, hourly rate, and contact information — all essential for transparency in grant-funded projects.

**Expenses Sheet**: This sheet links labor to specific site areas and materials, providing a clear record of costs for tasks such as masonry stabilization or roof replacement.

**Data Input Sheet**: Daily logging is simplified through dropdown menus and standardized job types, ensuring consistent recordkeeping across all restoration phases.

**Summary Sheet**: The summary aggregates payroll, expenses, and performance metrics, giving managers instant visibility into total labor costs, hours worked, and average rates.

---

For a restoration project like **Burrows Island Light Station**, this structured approach makes it easier to:

- Track qualified contractors  
- Manage federally or state-funded grants  
- Create accurate, audit-ready reports  

By introducing this level of precision and accountability, even small preservation teams can manage heritage projects efficiently while maintaining compliance with historical and financial documentation standards.

---

### Summary
This project integrates applied, Excel-based data management practices with the operational needs of heritage preservation. It shows how structured digital workflows enhance fiscal oversight, regulatory compliance, and sustainable management of cultural restoration initiatives.