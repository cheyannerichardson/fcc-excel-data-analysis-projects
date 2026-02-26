# Stanley Seawall Traffic Inventory

### Introduction

This beginner-friendly Excel project demonstrates how to build a dynamic car inventory system tailored for an urban planning traffic survey along the Stanley Park Seawall. The tool organizes vehicle data to support analysis of traffic patterns, congestion, and modal conflicts within this popular recreational corridor. The Stanley Park Seawall experiences frequent friction between pedestrians, cyclists, and vehicles accessing adjacent facilities, especially during peak hours and seasonal events. By applying clear scoring and tracking methods, this Excel framework helps planners evaluate patterns of use, identify safety hotspots, and guide design decisions that promote a more balanced, data-informed transportation environment.

## Project Overview

The **Stanley Seawall Traffic Inventory Project** is designed to create a structured and user-friendly Excel tool that captures, categorizes, and analyzes traffic data along the Stanley Park Seawall. The system supports urban planners and community stakeholders by compiling vehicle and mobility information to understand congestion trends, parking utilization, and safety concerns.

## Project Objectives

- Design an Excel-based traffic inventory to record and analyze non-vehicular mobility data.  
- Track pedestrian counts, cycling rates, micromobility usage, and adaptive mobility movements along the seawall.  
- Identify high-conflict zones between users, such as shared paths or intersections with limited visibility.  
- Support inclusive planning that improves mobility equity and user experience around the Stanley Park Seawall.
- Support evidence-based decision-making in the Stanley Park Seawall redevelopment and traffic management planning process.  
- Create a reusable template for future traffic studies and urban mobility assessments.

## Table of Contents

- [Introduction](#introduction)  
- [Project Overview](#project-overview)  
- [Project Objectives](#project-objectives)  
- [Data Collection Methodology](#data-collection-methodology)  
- [How to Build the Stanley Seawall Traffic Inventory Spreadsheet](#how-to-build-the-stanley-seawall-traffic-inventory-spreadsheet)  
- [Applying the System to Urban Planning](#applying-the-system-to-urban-planning)  
- [How to Use This Table for Your Own Projects](#how-to-use-this-table-for-your-own-projects)  
- [Conclusion](#conclusion)

## Data Collection Methodology

If an urban planner were creating the dataset for this project, they would begin by selecting survey points along key sections of the Stanley Park Seawall that represent different user conditions — such as high-traffic intersections, narrow shared pathways, scenic rest areas, and access points to nearby amenities. Survey teams would use consistent observation periods, typically covering peak and off-peak times across weekdays and weekends, to capture variation in user patterns. Data would be collected manually through tally sheets or digital counters, then entered into the Excel inventory using standardized categories for pedestrians, cyclists, micromobility devices, and adaptive mobility users. This structured approach ensures that measurements are comparable, reliable, and representative of real-world movement dynamics along the seawall.

### Select Survey Points

Hotspots (4 locations, 15-min intervals):
- Pipeline Road (vehicle pinchpoint)
- Second Beach (pedestrian heavy) 
- Ferguson Point (cyclist-ped conflicts)
- Causeway (all modes converge)

## How to Build the Stanley Seawall Traffic Inventory Spreadsheet

### Phase 1: Sheet Setup

1. **Create New File**  
   File → New → Blank Workbook → Save As **"SeawallTraffic_YYYYMMDD.xlsx"**

2. **Set Up Main Sheet**  
   Rename **Sheet1** to **"Seawall_Data"**  
   Add headers in **A1:H1**:  
   `ID | Type | Location | Count | Time | Direction | Status | Notes`

3. **Pre-Fill Dropdown Lists**  
   Use **Data → Data Validation**:

   - **Type (B:B):** Vehicle, Bike, Ped  
   - **Location (C:C):** Pipeline Rd, Second Beach, Causeway, Ferguson Pt  
   - **Direction (F:F):** CCW, CW  
   - **Status (G:G):** Normal, Slow, Conflict, Stopped

4. **Format for Field Use**  
   - View → Freeze Panes → **Freeze Top Row**  
   - Bold headers, yellow fill (Home → Fill Color)  
   - Auto-fit all columns  
   - Add **100 blank rows** for data entry  

5. **Convert Data to Table**  
   Select **A1:H100 → Insert → Table → Check “Headers” → OK**  
   Table Design → Rename table: **TrafficData**

6. **Add Visual Alerts (Conditional Formatting)**  
   For **Column G (Status):**

   - Select **G2:G100 → Home → Conditional Formatting → New Rule**  
   - Use formula: `=$G2="Conflict"`  
   - Format → **Yellow Fill → OK**
---

### Phase 2: Create Analysis Sheets

   **Sheet2: “Flow_Pivot”**  
   - Insert → PivotTable → New Worksheet  
   - Add fields:
     - **Rows:** Time (right-click → Group → Hours only → Format: h"PM")  
     - **Columns:** Type  
     - **Values:** Count (Sum)  
     - **Filters:** Location  
   - Insert → Slicer → Type (adds Bike/Ped toggle buttons)

   **Sheet3: “Conflicts_Pivots”**  
   - Rows: Location  
   - Columns: Type  
   - Filters: Status = Conflict only  
   - Values: Count (Sum)  
   - Insert → PivotChart → Clustered Column  

### Phase 3: Build Dashboard Summary

| Cell | Label / Formula | Description |
|------|------------------|--------------|
| A25 | **Total Conflicts** | Summary of all entries marked as "Conflict" |
| B25 | `=SUMIF(TrafficData[Status],"Conflict",TrafficData[Count])` | Calculates total conflicts |
| A26 | **Total Counts** | Total observed users (all types) |
| B26 | `=SUM(TrafficData[Count])` | Adds all counts in dataset |
| A27 | **Conflict Percentage** | Proportion of conflicts among total counts |
| B27 | `=B25/B26` | Calculates conflict rate (as a percent) |

**Tip:** Format **B27** as a percentage (Home → Number → %).

## Phase 4: Dashboard Creation and Reporting
 1. **Create the Chart**
   - With the PivotTable selected, go to **Insert → PivotChart → Clustered Column**.
   - This will generate a chart showing *locations* on the bottom (x-axis) and *number of conflicts* on the side (y-axis), grouped by *direction*.

2. **Customize Chart Appearance**
   - Add axis titles:
     - X-axis: “Location and Direction”
     - Y-axis: “Conflict Count”
   - Add chart title: **“Conflict Counts by Location and Direction”**

3. **Refine Visualization**
   - Ensure gridlines and legends are visible for clarity.
   - Format the legend labels clearly to match user types.
   - Resize and position the chart on your **Dashboard** worksheet for a balanced layout.

**Tip:**  
For presentation, right-click the chart → **Copy as Picture** → paste into a PowerPoint or PDF summary for visual reporting.

# Applying the System to Urban Planning
The Stanley Park Seawall serves as one of Vancouver’s most iconic recreational corridors, accommodating pedestrians, cyclists, adaptive mobility users, and micromobility devices in a fragile coastal environment. Its scenic design attracts both residents and tourists, but this popularity creates frequent congestion, directional conflicts, and accessibility concerns. Urban planners evaluating redevelopment options must therefore understand how, when, and where people move through the seawall network.

By applying the Excel-based traffic inventory system, planners can structure observational survey data similarly to a car inventory—organizing movements by type, time, location, and behavior. This standardization brings analytical clarity to complex mobility interactions. The dataset enables comparison across survey sites, highlighting which zones experience the highest ratios of conflicts or speed disparities. Over multiple survey periods, trends can reveal whether congestion is linked to infrastructure limitations, seasonal variation, or external influences such as event traffic or construction detours.

These insights support data-informed design decisions—such as where to widen pathways, add signage, or reconfigure lane separation. In long-term mobility planning, the inventory offers a repeatable foundation for performance monitoring, showing how interventions (e.g., new bollards, rerouted paths, or wayfinding improvements) affect user flow and safety outcomes.

Beyond Stanley Park, this structured Excel framework helps urban planners develop consistent methodologies for active transportation studies in other urban trails, plazas, and waterfront corridors. It institutionalizes evidence-based planning practices that align with broader goals of accessibility, sustainability, and equitable infrastructure investment—ensuring that each redesign not only enhances movement efficiency but also strengthens community experience and environmental stewardship.

# How to Use This Table for Your Own Projects

This traffic inventory table is built to be **flexible and customizable** for a wide range of urban mobility studies. Once the basic Excel structure is established—with headers for **Location**, **User Type**, **Direction**, **Count**, and **Status**—planners, students, or community groups can easily adapt the template to fit different contexts such as **neighborhood greenways**, **waterfront trails**, or **transit hub corridors**.

To use it, replace the sample site names with your actual **survey locations** and update the **dropdown lists** to reflect the modes of travel you’re tracking. Data collected during field surveys can be entered directly into the sheet, where **built-in formulas** automatically calculate totals and conflict percentages. 

The preconfigured **PivotTables and charts** instantly summarize flow rates, directional balance, and conflict intensity. By adjusting filters, users can visualize patterns by **time**, **location**, or **mobility type**, making this a practical and accessible tool for anyone conducting traffic or trail-use assessments—no specialized software required.

# Conclusion

The Stanley Seawall Traffic Inventory project demonstrates how a simple, structured Excel tool can serve as a powerful framework for collecting, analyzing, and visualizing active transportation data. By blending data management, urban design principles, and practical field methodology, this system empowers planners and community members to transform raw observations into actionable insights. Whether used to guide park redesigns, evaluate mobility equity, or support safety improvements, the approach promotes transparent, data-driven decision-making in urban trail and infrastructure planning. With minimal setup and freely available tools, anyone can replicate and refine this process to better understand—and improve—the flow of people and mobility across shared public spaces.
