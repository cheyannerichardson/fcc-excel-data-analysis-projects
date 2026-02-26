# 🐇 Tokkia Matcha House Sales Database

> **The Tokkia Matcha House Sales Database** is a beginner-friendly Excel project inspired by freeCodeCamp, designed to demonstrate how to build a fully integrated business management workbook. It organizes and connects key sales, cost, and procurement data for a real-world hospitality case study — the Tokkia Matcha House renovation and retail rollout. The project reflects practical challenges faced by small businesses, such as budget tracking, supplier coordination, and delivery delays. By bringing together financial, contractor, and procurement data into a single, dynamic dashboard, this system shows how structured spreadsheets can turn everyday business operations into clear, data-driven decisions.

## 🧩 Project Overview

The **Tokkia Matcha House Sales Database** is a structured Excel system designed to centralize cost tracking, contractor management, procurement, and performance reporting. It demonstrates practical spreadsheet automation concepts — from nested formulas and data validation to conditional formatting and chart-based dashboards.

### 🔍 Objectives
- Integrate all business data into a single, linked workbook.  
- Build clear financial visibility for budgeting and forecasting.  
- Simplify supplier and contractor management through cross-sheet references.  
- Create dynamic visuals and metrics that refresh automatically.  

## 📚 Table of Contents

1. [🧩 Project Overview](#-project-overview)
2. [📘 Project Database Setup (Step-by-Step)](#-project-database-setup-step-by-step)
   - [📄 Sheet 1: Cost & Budget Tracking](#-sheet-1-cost--budget-tracking)
   - [👷 Sheet 2: Contractor & Labor Management](#-sheet-2-contractor--labor-management)
   - [🚚 Sheet 3: Procurement & Deliveries](#-sheet-3-procurement--deliveries)
   - [📊 Sheet 4: Dashboard](#-sheet-4-dashboard)
3. [🏛️ Applying the System to Architecture](#-applying-the-system-to-architecture)
4. [🧭 How to Use This Table for Your Own Projects](#-how-to-use-this-table-for-your-own-projects)

##  Project Database Setup (Step-by-Step)

**Final Workbook Structure (4 Sheets):**
- **Cost & Budget Tracking** → Expenses & payment monitoring  
- **Contractor & Labor Management** → Team, contracts, and materials linkages  
- **Procurement & Deliveries** → Materials tracking and delivery insights  
- **Dashboard** → Executive summaries and charts  

---

## 📄 SHEET 1: COST & BUDGET TRACKING

### Step 1: Headers (Row 1)

A: Category | B: Item Description | C: Supplier | D: Quantity | E: Unit Cost (£)
F: Total Cost (£) | G: Budgeted Amount (£) | H: Variance (£) | I: Payment Status | J: Notes

text

---

### Step 2: Example Row

A2: Design
B2: Architectural Design Fees
C2: Stephenson-Edwards Studio
D2: 1
E2: 25000
F2: =D2*E2
G2: 28000
H2: =G2-F2
I2: Paid
J2: Final burrow layout approved Apr 5

text

👉 **Copy `F2` and `H2` formulas** down the table to auto-calculate totals and variances.

---

### Step 3: Data Validation

Range: I2:I100
Go to: Data → Data Validation → List
Options: "Paid, Pending, Partial (50%), Overdue"

text

---

### Step 4: Conditional Formatting

- **For Variance (H):** Format cells where value `< 0` → Red fill  

 Here’s your requested section written in clean, professional Markdown for your GitHub README so that it matches the style used in your earlier sections and renders beautifully:

text
## 👷 SHEET 2: CONTRACTOR & LABOR MANAGEMENT

### Step 1: Headers

A: Contractor | B: Role | C: Task | D: Start Date | E: End Date |
F: Contract Value (£) | G: Amount Paid (£) | H: Remaining (£) |
I: Status | J: Required Materials | K: Key Delivery Date | L: Delivery Status

text

---

### Step 2: Core Formulas

H2: =F2-G2

J2: =IF(A2="Lumina Studio","Pendant Lamps, Matcha Counter Lighting",
IF(A2="EcoBuild Ltd","Bamboo Floor Planks","General Materials"))

K2: =IF(ISNUMBER(SEARCH("Bamboo",J2)),
INDEX('Procurement and Deliveries'!D:D,
MATCH("Bamboo Floor Planks",'Procurement and Deliveries'!A:A,0)),"")

L2: =IF(J2="General Materials","",
IFERROR(INDEX('Procurement and Deliveries'!G:G,
MATCH(LEFT(TRIM(J2),FIND(",",TRIM(J2)&",")-1),
'Procurement and Deliveries'!A:A,0)),"No Match"))

text

---

## 🚚 SHEET 3: PROCUREMENT & DELIVERIES

### Step 1: Headers & Dropdowns

A: Item | B: Supplier | C: Order Date | D: Expected Delivery |
E: Received Date | F: Quantity | G: Status | H: Invoice #

text

**Create Status Dropdown:**
`G2:G100 → Data → Data Validation → List: "Pending, Delivered, Delivered (Late), Cancelled"`

---

### Step 2: Dashboard Summary (A23:B26)

A23: Total Items Ordered B23: =COUNTA(A2:A22)
A24: Pending Deliveries B24: =COUNTIF(G2:G22,"Pending")
A25: Delivered (Late) B25: =COUNTIF(G2:G22,"Delivered (Late)")
A26: % Delivered (on time) B26: =COUNTIF(G2:G22,"Delivered")/COUNTA(A2:A22)

text

---

### Step 3: Delivery Status Pie Chart

A28: Delivered B28: =COUNTIF(G2:G22,"Delivered")
A29: Pending B29: =COUNTIF(G2:G22,"Pending")
A30: Delivered Late B30: =COUNTIF(G2:G22,"Delivered (Late)")

text

**Insert:** Pie → 2-D Pie → Position near D23  
**Colors:**  
✅ Green = Delivered  
🟧 Yellow = Pending  
🔴 Red = Delivered Late  

---

## 📊 SHEET 4: DASHBOARD

### Step 1: Key Metrics (A1:B8)

A1: 🏗️ TOKKIA MATCHA HOUSE DASHBOARD (merge A1:F1)
A3: Total Budget B3: =SUM('Cost and Budget Tracking'!G:G)
A4: Total Spent B4: =SUM('Cost and Budget Tracking'!F:F)
A5: Remaining Budget B5: =B3-B4
A6: % Budget Spent B6: =B4/B3
A7: Over/Under Budget B7: =SUM('Cost and Budget Tracking'!H:H)
A8: Items Paid B8: =COUNTIF('Cost and Budget Tracking'!I:I,"Paid")

text

---

### Step 2: Top 5 Expenses

A11: TOP 5 EXPENSE CATEGORIES (merge A11:C11)
A12: Design B12: =SUMIF('Cost and Budget Tracking'!A:A,A12,'Cost and Budget Tracking'!F:F)
C12: =B12/$B$3 (copy down, format as %)

text

---

### Step 3: Budget Doughnut Chart

E20: Budget Remaining F20: =B3-B4
E21: Total Spent F21: =B4

text

Insert → **Doughnut Chart** → Colors:  
🟢 Green = Remaining  
🔴 Red = Spent  

---

### Step 4: Top-5 Bar Chart
Select `A12:B16` → Insert → Column Chart → Clustered Column  
Set title: **“Top 5 Expense Categories”**  
Enable **Data Labels → Outside End**

---

### Step 5: Procurement Summary

A20: Procurement Status
A21: Pending Deliveries B21: =COUNTIF('Procurement and Deliveries'!G:G,"Pending")
A22: Delivered On Time B22: =COUNTIF('Procurement and Deliveries'!G:G,"Delivered")
A23: Delivered Late B23: =COUNTIF('Procurement and Deliveries'!G:G,"Delivered (Late)")

---

## 🎨 Formatting & Branding
- **Currency:** Home → Format Cells → Currency (£)  
- **Percentages:** 0 decimal points  
- **Dashboard:** View → Freeze Top Row  
- **Conditional Formatting:**  
  - Variance < 0 → 🔴 Red fill  
  - Pending → 🟧 Orange fill  
  - Delivered → 🟩 Green fill  
- **Charts:** Use *earthy green & brown tones* for a Tokkia brand aesthetic  

---

## ✅ Final Deliverable
You’ll have a **4-tab integrated Excel workbook** with:  
- 📈 Auto-updating totals and variances  
- 🔄 Linked contractor and procurement data  
- 📊 Real-time dashboard metrics  
- 🎯 Executive-ready visual summaries  

## 🏛️ Applying the System to Architecture

Although the Tokkia Matcha House Sales Database was developed for a café-based business, its structure directly supports the way architecture and design firms plan, manage, and deliver their projects. Architectural practices routinely balance multiple cost categories—consultant fees, materials procurement, labor contracts, and client billing—making this system a valuable, data-driven project management framework. By consolidating all of these operations into a single workbook, firms can create a transparent record of budgets, timelines, and supplier performance in real time. 

The database functions like a lightweight digital project management hub, ideal for studios handling hospitality, retail, or interior architecture projects where precision and coordination are critical. Each linked sheet mirrors how architectural workflows evolve: from conceptual design budgets and contractor tenders to material sourcing and delivery tracking. As a result, team leads gain immediate insights into spending patterns, delivery delays, and payment variances—allowing proactive decision-making instead of reactive problem-solving.  

For creative and boutique studios, this approach bridges the gap between artistic design and operational management, embedding financial clarity within the design process itself. The system effectively transforms ordinary spreadsheet tools into an architectural intelligence dashboard—aligning design intent, cost control, and project delivery under a single, accessible platform.

## 🧭 How to Use This Table for Your Own Projects

To adapt the table for your professional workflow, begin by replacing the sample data in each sheet with your own categories, tasks, and suppliers.  
- In **Sheet 1 (Cost & Budget)**, define your cost centers—such as design, materials, or fabrication—and enter associated unit costs and quantities.  
- In **Sheet 2 (Contractor & Labor)**, list your key consultants, contractors, and deliverables, linking their required materials and start/end dates.  
- In **Sheet 3 (Procurement & Deliveries)**, connect ordered items with supplier data to track delivery performance automatically.  
- The **Dashboard** will then update live charts and KPIs to visualize your budget efficiency and procurement reliability.

Professionals can use this table not just as a budget ledger, but as a dynamic project intelligence tool—translating individual line items into a clear narrative of resource use, workflow efficiency, and delivery accountability across the entire project.

### 🧩 Conclusion

The Tokkia Matcha House Sales Database demonstrates how an adaptable spreadsheet system can empower architecture and design firms to work more strategically. By integrating financial data, contractor performance, and procurement insights, the platform encourages data-informed creativity—enabling teams to balance aesthetics with accountability. Whether applied to urban renovations, residential interiors, or multi-phase commercial builds, this model promotes transparency, efficiency, and professional polish in every project workflow.
