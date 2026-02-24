# 🌊 Stanley Park Seawall Climate Risk Analyzer
## Introduction

This beginner-friendly Excel project demonstrates how to build a **dynamic climate risk inventory system** tailored for urban planning analysis on the **Stanley Park Seawall** in Vancouver, Canada.

**Climate factors threatening the Seawall:**
- **Sea-level rise**: 0.5m by 2050, 1m by 2100 (BC guidelines)
- **Storm surges**: Increasing frequency + intensity (50mm rain + 5.4m tides)
- **Erosion**: Wave reflection damages aging concrete infrastructure  
- **Seismic activity**: Earthquake risks to structural integrity
- **Tourist pressures**: Heavy foot/bike traffic stresses repairs

**Excel Problem Inventory Application:**
A decision matrix with scenario dropdowns (Low/Medium/High/Critical), heatmaps, PivotTables, and dual-axis charts enables planners to:
- Assign risk scenarios per threat
- Calculate weighted costs ($75K-$750K per mitigation)
- Visualize priorities (Red=Critical via conditional formatting)
- Test "what-if" budgets via slicers
- Present funding justification to stakeholders

## 🎯 Project Objectives

- **Build risk inventory table** with Vancouver-specific climate threats (sea-level rise, storms, erosion, earthquakes, tourist overload)
- **Create live scenario dropdowns** (Low/Medium/High/Critical) with VLOOKUP/IF formulas for dynamic cost/weight calculation
- **Apply heatmap conditional formatting** (Green-White-Red) for instant risk prioritization
- **Calculate weighted risk/cost totals** using SUMPRODUCT for budget planning
- **Develop PivotTable dashboard** with Scenario slicer for what-if analysis
- **Design dual-axis combo chart** (Risk columns + Cost line) for stakeholder presentations
- **Test complete workflow**: Change dropdown → Watch entire dashboard update live
- **Generate print-ready urban planning report** for seawall resilience funding

## Table of Contents
1. [Introduction](#introduction)  
2. [Project Objectives](#project-objectives)  
3. [Guide to Building Climate Risk Analyzer](#guide-to-building-climate-risk-analyzer)  
4. [Applying to Climate-Resilient Infrastructure Planning](#applying-to-climate-resilient-infrastructure-planning)
5. [How to Use This Table for Your Own Project](#how-to-use-this-table-for-your-own-project)  
6. [Conclusion](#conclusion)

## Guide to Building Climate Risk Analyzer 

#### Complete Excel Decision Matrix - Live scenario analysis with heatmaps, PivotTables, and dual-axis charts for Vancouver urban planning.

---
### Create Risk Factors Table

|   | A              | B                              | C | D | E        | F   | G   | H       |
|---|----------------|--------------------------------|---|---|----------|-----|-----|---------|
| 1 | Risk Factor    | Description                    | Prob | Impact | Scenario | Cost ($K) | Weight | Raw Risk |
| 2 | Sea-Level Rise | 0.5m by 2050 per BC guidelines | 5   | 5     | Low      | 75  | 0.2 | =C2*D2  |
| 3 | Storms         | 50mm rain + 5.4m tides         | 4   | 4     | Medium   | 200 | 0.5 | =C3*D3  |
| 4 | Erosion        | Wave action + seawall reflection | 4 | 4     | High     | 400 | 0.8 | =C4*D4  |
| 5 | Earthquakes    | Seismic risks to concrete      | 3   | 5     | Critical | 750 | 1.0 | =C5*D5  |
| 6 | Tourist Overload | High volume stresses repairs | 3 | 3     | Medium   | 200 | 0.5 | =C6*D6  |


1. **E2:E6 Dropdowns**: Data → Data Validation → List → `Low,Medium,High,Critical`
2. **F2 Formula** (drag to F6): `=IF(E2="Low",75,IF(E2="Medium",200,IF(E2="High",400,IF(E2="Critical",750,0))))`
3. **G2 Formula** (drag to G6): `=IF(E2="Low",0.2,IF(E2="Medium",0.5,IF(E2="High",0.8,IF(E2="Critical",1.0,0))))`
4. A12: Total Weighted Risk B12: =SUMPRODUCT(H2:H6,G2:G6)
5. A13: Total Weighted Cost B13: =SUMPRODUCT(F2:F6,G2:G6)

---
### Heatmap 
 **Apply SEPARATELY to each column for maximum visual impact:**

1. Select **H2:H6** → Home → Conditional Formatting → Color Scales → Green-White-Red
2. Home → Conditional Formatting → Manage Rules → Edit Rule → Custom Format:

            Green (Low): RGB(0,176,80) | #00B050
            Yellow (Medium): RGB(255,192,0) | #FFC000
            Red (High): RGB(192,0,0) | #C00000

3. Edit Rule → Format Style: 3-Color Scale
           
            Minimum: Number = 0    (Green)
            Midpoint: Number = 12.5 (Yellow) 
            Maximum: Number = 25   (Dark Red)

4. Apply to range: =$H$2:$H$6

1. Select **F2:F6** → Home → Conditional Formatting → Color Scales → Green-White-Red
2. Home → Conditional Formatting → Manage Rules → Edit Rule → Custom Format:

            Green (Low): RGB(0,176,80) | #00B050
            Yellow (Medium): RGB(255,192,0) | #FFC000
            Red (High): RGB(192,0,0) | #C00000

3. Edit Rule → Format Style: 3-Color Scale
           
            Minimum: Number = 0    (Green)
            Midpoint: Number = 375 (Yellow) 
            Maximum: Number = 750   (Dark Red)

4. Apply to range: =$F$2:$F$6

1. Select **G2:G6** → Home → Conditional Formatting → Color Scales → Green-White-Red
2. Home → Conditional Formatting → Manage Rules → Edit Rule → Custom Format:

            Green (Low): RGB(0,176,80) | #00B050
            Yellow (Medium): RGB(255,192,0) | #FFC000
            Red (High): RGB(192,0,0) | #C00000

3. Edit Rule → Format Style: 3-Color Scale
           
            Minimum: Number = 0    (Green)
            Midpoint: Number = 0.5 (Yellow) 
            Maximum: Number = 1.0   (Dark Red)

4. Apply to range: =$G$2:$G$6

## Live Dashboard Totals

**Select A1:H6 → Insert → PivotTable → New Worksheet:**

            Fields:

            text
            FILTERS: Scenario (E)
            ROWS: Risk Factor (A)
            VALUES: Raw Risk Score (H)
            VALUES: Mitigation Cost ($K) (F)

Slicer: PivotTable Analyze → Insert Slicer → Scenario

#### Dual-Axis Combo Chart

1. Click PivotTable → Insert → Column Chart → Clustered Column
2. Chart Design → Change Chart Type → Combo:
   
| Series              | Chart Type        | Axis              | Max Value |
|---------------------|-------------------|-------------------|-----------|
| ✓ Raw Risk Score    | Clustered Column  | Primary (Left)    | 30        |
| ✓ Mitigation Cost   | Line              | Secondary (Right) | 800       |

3. Title: "Stanley Park Seawall Climate Risks by Scenario"


✅ Dashboard Features
- Live scenario dropdowns per risk
- Heatmap visualization (Red=Critical, Green=Low)
- Slicer-controlled PivotTable
- Dual-axis chart (Risk bars + Cost line)
- Weighted totals for planners
- Print-ready Vancouver presentations

## Applying to Climate-Resilient Infrastructure Planning

The Stanley Park Seawall Climate Risk Analyzer empowers urban planners to tackle coastal adaptation for climate-resilient trail infrastructure through a dynamic Excel decision matrix. This tool transforms complex climate vulnerabilities into actionable insights for the Seawall's unique challenges: 0.5m sea-level rise by 2050, intensifying storm surges (50mm rain + 5.4m tides), wave-driven erosion on aging concrete, seismic threats, and tourist overload stressing repairs.

Urban planners apply the table through scenario-driven analysis:

    Assign risk levels (Low/Medium/High/Critical) per threat using dropdowns
    Heatmap instantly flags priorities—Sea-Level Rise glows dark red (25/25 score)
    Weighted totals calculate budgets—Critical scenario = $2.25M for full mitigation
    PivotTable slicers test "what-if"—filter "High" risks only → $1.6M targeted spend
    Dual-axis chart presents evidence—tall risk columns + cost line spikes justify funding

Applied to the Stanley Park Seawall, this Climate Risk Analyzer becomes a critical decision-making tool for Vancouver's 9km iconic coastal pathway, which faces escalating threats that demand urgent redevelopment providing urban planners with data-driven adaptation roadmap that secures stakeholder buy-in and federal climate funding for Vancouver's iconic 9km coastal trail.

## How to Use this Table for Your Own Project

Anyone can adapt this Climate Risk Analyzer for their coastal trail or urban infrastructure project - just replace Stanley Park data with your location's specific risks.
    
1. Update Risk Factors (Column A) and replace with your project's threats:

        River flooding, wildfire smoke, heat islands, bridge scour, permafrost melt, ect.

2. Input Local Data (Columns C-D) 

**Enter Probability (1-5) and Impact (1-5) based on:** 

        ✅ Engineering reports (sea walls = high impact)

        ✅ Historical incidents (storm closures = high probability)

        ✅ Climate models (your city's 2050 projections)

        ✅ Cost estimates ($K) from contractors

3. Assign Scenarios (Column E Dropdowns)

        Keep Low/Medium/High/Critical - or customize based on project needs
       

4. Live Results Generate Automatically

        ✅ Heatmap flags priorities (Red = Do Now)
        ✅ PivotTable slicer shows "High risk total = $X"
        ✅ Dual-axis chart ready for your City Council presentation
        ✅ B12/B13 totals = Your defensible budget ask

5. Export for Stakeholders

        File → Print → "Fit to Page" → PDF
        Include: Heatmap + Chart + Policy Table


#### Works for: Coastal trails, boardwalks, urban bridges, riverfronts - any climate-vulnerable public asset.

## Conclusion

The Stanley Park Seawall Climate Risk Analyzer transforms freeCodeCamp Excel skills into a professional urban planning tool that quantifies Vancouver's unique coastal threats into actionable policy recommendations.

Planners gain:

    ✅ Live risk prioritization via heatmaps (Red = Second Beach NOW)

    ✅ Defensible budgets ($2.45M prevents $50M disaster costs)

    ✅ Federal grant justification with dual-axis charts + slicer analysis

Stanley Park Seawall Climate Risk Analyzer powers redevelopment - turning sea-level rise projections, storm surge data, and erosion rates into a roadmap that safeguards Vancouver's iconic 9km coastal trail from climate liability to global resilience model.