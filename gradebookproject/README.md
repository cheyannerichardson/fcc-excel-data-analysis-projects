# Stanley Park Seawall Community Feedback & Impact Gradebook

A freeCodeCamp beginner-friendly project demonstrating how to build a weighted gradebook in Excel to organize and evaluate community feedback and impact metrics for the Stanley Park Seawall redevelopment project using real-world environmental, social, and design data. The **Stanley Park Seawall** faces complex challenges, from erosion and sea-level rise to balancing recreation and habitat protection. This project delivers a structured Excel system that helps stakeholders evaluate multiple design proposals using a clear and replicable scoring method.

---

## Project Overview

### Project Objectives
- Enable **transparent and consistent evaluation**  
- Reflect **community and policy priorities** through weighted metrics  
- Provide **automated rankings and curated visuals** to support informed decision-making  

---

## 🗂️ Table of Contents

1. [Project Overview](#-project-overview)  
2. [Proposal Descriptions](#-proposal-descriptions)  
3. [Test Data Setup](#-test-data-setup)   
7. [Understanding the Scores and Ranking](#-understanding-the-scores-and-ranking)  
8. [Applying the System to Urban Planning](#-applying-the-system-to-urban-planning)  
9. [How to Use This Table for Your Own Projects](#-how-to-use-this-table-for-your-own-projects)  
10. [Next Steps for Urban Planners](#-next-steps-for-urban-planners)

---

## 🌊 Proposal Descriptions

Each proposal in the Excel gradebook represents a distinct vision for restoring and reimagining the Stanley Park Seawall.

### 1. **Living Shoreline Reconnection**
Emphasizes **ecological restoration** and **intertidal biodiversity** by replacing hard seawall edges with soft shoreline terraces and native vegetation.  
- **Focus:** Climate resilience, habitat restoration, and immersive visitor experience.  
- **Tradeoff:** May slightly slow cycling or pedestrian movement near wildlife areas.  

---

### 2. **Layered Mobility Path System**
Creates dedicated **pedestrian, cyclist, and maintenance paths** for safer, more efficient use of the waterfront.  
- **Focus:** Inclusive design, traffic efficiency, and recreation flow.  
- **Tradeoff:** Expands construction footprint, slightly reducing natural shoreline area.  

---

### 3. **Eco-Art Storyline Walk**
Merges **Indigenous art, interpretive signage, and ecological storytelling** along restored habitats to celebrate coastal heritage.  
- **Focus:** Cultural education, inclusivity, and experiential learning.  
- **Tradeoff:** Higher maintenance demands and potential for localized crowding.  

---

### 4. **Tidal Amphitheatre and Learning Cove**
Develops amphitheater-style terraces at select sites for tidal observation, outdoor classrooms, and performances.  
- **Focus:** Recreation, education, and environmental awareness.  
- **Tradeoff:** Moderate disruption during construction; needs erosion control measures.  

---

### 5. **Adaptive Elevation Pathway**
Elevates vulnerable segments of the seawall to counter **sea-level rise** and **storm damage**, integrating vegetated slopes for stability.  
- **Focus:** Climate resilience, long-term safety, and structural durability.  
- **Tradeoff:** Visual changes and temporary detours during installation.  

---

### 6. **Wildlife Passage and Observation Zones**
Designates **buffer zones** and **viewing decks** to protect sensitive habitats while maintaining visitor access.  
- **Focus:** Biodiversity preservation and ecological stewardship.  
- **Tradeoff:** Reduces direct access to some scenic waterfront sections.  

---

### 7. **Heritage Continuity Loop**
Conserves **historic stonework** while incorporating new materials and grades for accessibility improvements.  
- **Focus:** Heritage preservation, visitor comfort, and low-impact adaptation.  
- **Tradeoff:** Limited adaptability for large-scale ecological redesign.  

---

## ⚙️ System Setup and Instructions

### Core Functions
- Input and weight **impact metrics** (total weight = 1.00)  
- Evaluate **7 proposals (rows)** across **11 metrics (columns B–L)**  
- Automatically generate **averages, weighted scores, and rankings**  
- Use built-in **color-coded visuals and top-choice output**

## 🧮 Test Data Setup

Enter the following data into your Excel sheet (Cells **A1:L8**).  
Each row contains evaluation scores (1–5) for the 7 seawall proposals across 11 metrics (columns **B–L**).

| Proposal | B | C | D | E | F | G | H | I | J | K | L |
|-----------|---|---|---|---|---|---|---|---|---|---|---|
| **A2: Living Shoreline** | 5 | 5 | 4 | 5 | 3 | 4 | 4 | 5 | 4 | 3 | 4 |
| **A3: Layered Mobility** | 4 | 3 | 3 | 4 | 5 | 5 | 3 | 4 | 4 | 5 | 5 |
| **A4: Eco-Art** | 4 | 4 | 3 | 3 | 4 | 5 | 5 | 5 | 4 | 3 | 4 |
| **A5: Tidal Amphitheatre** | 4 | 4 | 4 | 4 | 4 | 4 | 4 | 5 | 4 | 4 | 4 |
| **A6: Adaptive Elevation** | 5 | 4 | 5 | 5 | 3 | 3 | 3 | 4 | 4 | 3 | 4 |
| **A7: Wildlife Passage** | 5 | 5 | 4 | 5 | 2 | 3 | 3 | 4 | 3 | 2 | 3 |
| **A8: Heritage Loop** | 3 | 3 | 3 | 4 | 4 | 4 | 5 | 4 | 4 | 5 | 5 |

---

## 🧭 [Step 3: Create Weights Row (Row 20)]()

### 3.1 Label & Enter Weights
| Cell | Entry |
|-------|--------|
| **A20** | `Metric Weights (Total = 1.00)` *(Bold, 12pt)* |
| **B20–L20** | `0.15, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10, 0.10, 0.05, 0.05, 0.05` |

### 3.2 Verify Total Weight
| Cell | Formula | Expected Result |
|-------|----------|----------------|
| **M20** | `=SUM(B20:L20)` | `1.00` ✅ |

---

## ⚙️ [Step 4: Build Scoring Formulas]()

### 4.1 Column M – **Average Score**

| Cell | Formula |
|-------|----------|
| `M2` | `=AVERAGE(B2:L2)` |

Copy **M2** down to **M8**.

---

### 4.2 Column N – **Weighted Score**

| Cell | Formula |
|-------|----------|
| `N2` | `=SUMPRODUCT(B2:L2,$B$20:$L$20)` |

Copy **N2** down to **N8**.

---

### 4.3 Column O – **Sequential Rank (No Ties)**

| Cell | Formula |
|-------|----------|
| `O2` | `=IFERROR(RANK(N2,$N$2:$N$8,0),"Check Data")` |

Copy **O2** down to **O8**.  
This ensures **unique ranking** even if two weighted scores are identical.

---

## [Step 5:📎 File Formatting Summary

| Column | Content | Format |
|---------|----------|--------|
| A | Proposal names | Text |
| B–L | Metric scores (1–5) | Integer |
| M | Average Score | Number (2 decimals) |
| N | Weighted Score | Number (2 decimals) |
| O | Rank | Integer |

- **Columns M & N:** 2 decimal places (Home → Increase Decimal → 2).  
- **Column O:** Whole number format.

---

## 🏆 [Step 6: Create Decision Summary (Row 12)]()

### 6.1 Top Choices Display
| Cell | Formula / Label |
|-------|-----------------|
| **A12** | `TOP RECOMMENDATIONS` *(Bold, 14pt)* |
| **A13** | `#1 Choice:` |
| **B13** | `=INDEX(A$2:A$8,MATCH(1,O$2:O$8,0))` |
| **C13** | `Weighted Score:` |
| **D13** | `=INDEX(N$2:N$8,MATCH(1,O$2:O$8,0))` |
| **A14** | `#2 Choice:` |
| **B14** | `=INDEX(A$2:A$8,MATCH(2,O$2:O$8,0))` |

This section dynamically displays the highest-ranked projects and their corresponding weighted scores.

## Project Rankings
![Clustered Column Chart](chart-screenshot.png)
*Side-by-side comparison of 8 projects across 12 ranking criteria*

## Project Proposal Rankings

Create a **Clustered Bar Chart** to compare project proposals:

### Step-by-Step (Excel)

1. **Data Setup:**

A2:A8 = Project proposal names
O2:O8 = Rank numbers (1 = best)

text

2. **Insert Chart:**

Select A1:O8 → Insert → Charts → Clustered Bar (2-D Bar section)

text

3. **Fix Data Series:**

Right-click chart → Select Data:
• Series Name: Rank
• Series Values: ='Stanley Park Seawall'!$O$2:$O$8
• Horizontal Axis: ='Stanley Park Seawall'!$A$2:$A$8

text

4. **Format:**

• Reverse vertical axis (Rank 1 at top)
• Tilt project names -45° for readability

text

### Result
Horizontal bars showing rank lengths, project names on left axis. Perfect for proposal evaluation dashboard.

![Project Rankings Bar Chart](chart-screenshot.png)
*Top-ranked proposals have longest bars*
 
Step 8: Final Validation Checklist
✅ All these must be true:
[ ] M20 shows: 1.00 (weights total)
[ ] Column M: 3.8-4.3 range (averages)
[ ] Column N: 3.8-4.3 range (weighted)
[ ] Column O: 1,2,3,4,5,6,7 (sequential ranks)

---

## ## Understanding the Scores and Rankings

The spreadsheet uses three core comparative values to evaluate each seawall proposal.

- ### **Average Score**
- Represents the **general performance** across all metrics.
- Treats all impact areas equally. Overall performance across all metrics (unweighted).  
- **Weighted Score:** Reflects a proposal’s performance aligned with strategic priorities.
Multiplies each metric’s score by its assigned weight (importance) before totaling.
Gives higher influence to the criteria that matter most, such as resilience or community benefit.
  
- **Rank:** Final order used in the summary chart and top-choice display.

After entering scores and formulas, Excel will:
- Calculate **Average & Weighted Scores** for each proposal.  
- Automatically rank all seven proposals based on total weighted score.  

---

***

## 🧱 Application and Use Cases

### Applying the System to Urban Planning  

The **Stanley Park Seawall Restoration Gradebook** is a practical example of how **urban planners** can use structured evaluation tools to align technical decisions with community values.  
For the seawall project, stakeholders such as residents, environmental groups, accessibility advocates, and city engineers each bring unique priorities—ranging from **habitat restoration and mobility safety** to **heritage conservation and climate resilience**.  

This gradebook converts those priorities into measurable data using **weighted impact metrics**, which are developed from **community feedback surveys** and **stakeholder consultations**.  
Urban planners can assign greater weights to the criteria most valued by respondents (for instance, climate resilience or public access). When scoring each design proposal—like the *Living Shoreline Reconnection* or *Adaptive Elevation Pathway*—the system automatically generates **average, weighted, and ranked results** that reveal which designs most effectively align with the collective goals identified through consultation.  

In practice, this process strengthens transparency and trust in public decision-making. The ranked tables make it straightforward to communicate why certain proposals rise to the top, helping planners present clear evidence of how public input shaped final priorities.  
For the **Stanley Park Seawall**, this method ensures that restoration strategies are not only technically robust but also **socially inclusive and environmentally responsive**—a model that other cities can replicate for **coastal adaptation**, **urban trail systems**, or **climate‑resilient infrastructure planning**.

---

### How to Use This Table for Your Own Projects  

This gradebook can be adapted to almost any **urban planning or infrastructure** project that combines technical criteria with public engagement.  
Begin by identifying the **key evaluation metrics** that matter for your site—such as sustainability, cost, aesthetic value, accessibility, cultural impact, or climate resilience.  
Collect **community or stakeholder feedback** through surveys, interviews, or workshops to set **weights** that reflect their priorities.  

Next, input evaluation scores for each design option or contractor across the selected metrics. The workbook will automatically calculate **average**, **weighted**, and **ranked scores**, offering a transparent overview of trade‑offs and strengths.  
Visual dashboards—such as bar or radar charts—can then illustrate which proposals best meet collective objectives.  

Whether applied to **coastal engineering**, **parks and open‑space redesign**, or **urban renewal**, this approach enables planners to merge **data‑driven analysis** with **community voice**, ensuring that final decisions are both accountable and visionary—just as demonstrated by the **Stanley Park Seawall Restoration project**.

### Next Steps for Urban Planners  

For **urban planners**, this gradebook is the foundation of an evidence‑based planning process that connects community input, design evaluation, and data management.  
After using Excel to score and rank proposals, planners can integrate the results into larger **planning workflows** to guide design refinement, funding decisions, and policy reporting.  

**1. Data Management and Analysis**  
Export the gradebook results to tools like **Microsoft Access** or **SQL databases** to store and compare multiple evaluation cycles or project alternatives. This allows planners to track how designs evolve and which criteria consistently drive community preference or policy alignment.  

**2. Spatial Visualization and Mapping**  
Link proposal data with spatial layers using **GIS software** such as **ArcGIS**, **QGIS**, or **MapInfo**. This step visualizes where each proposal delivers benefits—like improved flood protection, shoreline access, or ecological restoration—helping planners communicate spatial trade‑offs across the site or region.  

**3. Advanced Metrics and Modeling**  
For large‑scale infrastructure or multi‑year planning initiatives, integrate data from environmental models, traffic simulations, or climate‑risk tools into the database. By doing so, the planner creates a **dynamic feedback system** that connects performance metrics from the gradebook directly to spatial and temporal analysis.  

**4. Implementation Tracking and Reporting**  
The weighted‑ranking structure of this gradebook also supports **progress monitoring** once implementation begins. As updates are made, planners can refresh scores or weights to reflect new information and visualize project performance over time through dashboards or GIS‑linked web maps.  

---

In the context of the **Stanley Park Seawall Restoration**, these next steps turn the gradebook from a decision‑making table into a **comprehensive planning platform**.  
By linking Excel’s scoring transparency with the analytical power of **GIS, databases, and visualization software**, urban planners can not only compare design options but also map, monitor, and communicate outcomes—ensuring that each recommendation remains both technically sound and publicly accountable.
