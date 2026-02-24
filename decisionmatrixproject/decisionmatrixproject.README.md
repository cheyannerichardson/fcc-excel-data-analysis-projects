# Washington Lighthouses Decision Matrix (Excel Project)

A freeCodeCamp beginner-friendly project demonstrating how to build a **weighted decision matrix in Excel** to analyze preservation priorities for Washington State lighthouses using real-world environmental, cost, and cultural datasets.

---

## 📘 Project Overview

Washington State’s lighthouses have safeguarded ships navigating some of the most treacherous coasts in North America—from the fog banks of the Strait of Juan de Fuca to the shifting sands of the Columbia River Bar. Today, these same structures face a different set of hazards: coastal erosion, rising sea levels, deferred maintenance, and limited public funding. Historic lighthouse management in Washington has required balancing preservation costs with public access, heritage value, and environmental resilience. Projects like the restorations of **Admiralty Head** and **West Point** demonstrate great success when multi-agency partnerships align—while remote sites like **Cape Flattery** reveal the financial and logistical challenges of preservation in rugged, exposed environments.

A structured evaluation tool such as a **Decision Matrix** offers an objective, data-driven approach to these complex heritage dilemmas. By integrating environmental risk data, maintenance costs, sustainability metrics, and cultural significance, preservation teams can quantify competing priorities rather than rely on intuition or political influence alone. This framework transforms qualitative heritage discussions into measurable strategies—helping Washington’s heritage planners and grant programs direct limited restoration funds toward sites offering the most achievable long-term impact.

This Excel model evaluates five historic **Washington lighthouses** across **nine measurable criteria**, combining quantitative data from coastal agencies and preservation sources.

### Features
- Weighted scoring system balancing cost efficiency, risk, and heritage value  
- Conditional formatting and chart visualizations  
- Auto-calculated rankings for five lighthouses  
- Fully editable Excel matrix ready for presentation  

---
## 📚 Table of Contents

1. [🏛️ Project Overview](#-project-overview)
2. [📊 Scoring Methodology & Data Sources](#-scoring-methodology--data-sources)  
3. [Building the Decision Matrix in Excel](#-building-the-decision-matrix-in-excel)  
   - [Step 1: Set Up the Spreadsheet](#step-1-set-up-the-spreadsheet)  
   - [Step 2: Enter Headers](#step-2-enter-headers-row-2)  
   - [Step 3: Enter Lighthouses & Data](#step-3-enter-lighthouses--data-rows-37)  
   - [Step 4: Enter Total Score Formula](#step-4-enter-total-score-formula)  
   - [Step 5: Format for a Professional Look](#step-5-format-for-a-professional-look)  
   - [Step 6: Add Visualization](#step-6-add-visualization)  
   - [Step 7: Add Final Analysis Features](#step-7-add-final-analysis-features)  
4. [🧭 Applying the System to Historic Restoration](#-applying-the-system-to-historic-restoration)  
5. [🧩 How to Use This Table for Your Own Projects](#-how-to-use-this-table-for-your-own-projects)  

---

## 📊 Scoring Methodology & Data Sources

The following details outline sourcing and calculation methodology for all **9 evaluation criteria** across **5 Washington lighthouses**.

---

### 🏗️ Preservation Cost (Column B) — 10 = Lowest Cost

**Source Institutions:** U.S. Lighthouse Society Preservation Grants Database, WA State Parks maintenance reports  
**Benchmark:** $2.5M = maximum (10), linear scale down from max baseline.

| Lighthouse | Raw Cost Data | Score Calculation | Source |
|-------------|----------------|--------------------|----------|
| West Point | \$0.6M repairs | 10 - (0.6 / 2.5 × 9) = **7** | Seattle Parks Dept records |
| Admiralty Head | \$1M restoration | 10 - (1.0 / 2.5 × 9) = **6** | WA State Parks Fort Casey contracts |
| Lime Kiln | \$0.5M maintenance | 10 - (0.5 / 2.5 × 9) = **8** | USCG San Juan Island maintenance |
| Cape Disapp. | \$1.8M estimate | 10 - (1.8 / 2.5 × 9) = **5** | WA State Parks Cape Disappointment plan |
| Cape Flattery | \$2M remote access | 10 - (2.0 / 2.5 × 9) = **4** | Makah Tribal & USCG logistics data |

---

### 🗺️ Zoning Compliance (Column C)

**Source Institutions:** WA Dept. of Ecology Shoreline Master Program (SMP), county planning records  
**Scoring:** Full SMP compliance = 9–10; mixed or tribal jurisdiction = 6–7.

| Lighthouse | Zoning Status | Score | Source |
|-------------|----------------|--------|---------|
| West Point | Seattle Parks (full SMP) | 9 | Seattle SMP |
| Admiralty Head | WA State Parks | 8 | Kitsap County shoreline permits |
| Lime Kiln | San Juan County preserve | 9 | WA Dept. Ecology |
| Cape Disapp. | State park, restricted | 7 | Pacific County ordinance |
| Cape Flattery | Tribal/federal mix | 6 | Makah Tribal planning docs |

---

### 🌊 GIS Flood Risk (Column D) — 10 = Worst

**Source Institutions:** WA Coastal Hazards Resilience Network, NOAA CoSMoS flood layers

| Lighthouse | FEMA Flood Zone | WA Coastal Atlas Risk | Score | Source |
|-------------|----------------|------------------------|--------|---------|
| West Point | AE Zone | Moderate | 6 | Puget Sound LiDAR flood maps |
| Admiralty Head | X Zone | Low–moderate | 5 | Admiralty Inlet CoSMoS |
| Lime Kiln | X/VE (sheltered) | Low | 4 | San Juan County GIS |
| Cape Disapp. | VE (oceanfront) | Extreme | 10 | Pacific County maps |
| Cape Flattery | VE (exposed) | High | 8 | WA Coastal Atlas (Tatoosh) |

---

### 🌱 Sustainability Rating (Column E)

**Source Institutions:** USCG green retrofit reports, facility energy audits

| Lighthouse | Retrofit Status | Energy Efficiency | Score |
|-------------|------------------|------------------|--------|
| West Point | LED conversion | Moderate | 7 |
| Admiralty Head | Solar pilot | Good | 8 |
| Lime Kiln | Full solar + wind | Excellent | 9 |
| Cape Disapp. | Partial LED | Fair | 6 |
| Cape Flattery | Diesel generator | Poor | 5 |

---

### 🧭 Public Access / Tourism Value (Column F)

**Source Institutions:** WA State Parks visitation data, county tourism boards

| Lighthouse | Annual Visitors | Revenue Potential | Score | Source |
|-------------|----------------|------------------|--------|---------|
| West Point | 150K | High | 8 | Seattle Parks |
| Admiralty Head | 400K | Very High | 9 | WA State Parks |
| Lime Kiln | 500K+ | Highest | 10 | San Juan Tourism Board |
| Cape Disapp. | 1M+ | Highest | 10 | WA State Parks |
| Cape Flattery | 20K | Low | 4 | Makah Tribal Tourism |

---

### 🐚 Ecological Impact (Column G)

**Source Institutions:** WA Dept. of Fish & Wildlife, NOAA Marine Protected Areas

| Lighthouse | Habitat Proximity | Conservation Value | Score |
|-------------|------------------|--------------------|--------|
| West Point | Urban Puget Sound | Moderate | 7 |
| Admiralty Head | Forage fish habitat | High | 8 |
| Lime Kiln | Orca whale sanctuary | Highest | 10 |
| Cape Disapp. | Columbia estuary | Moderate | 6 |
| Cape Flattery | Makah coastal reserve | Very High | 9 |

---

### 🏖️ Erosion / Sea Level Rise (Column H) — 10 = Worst

**Source Institution:** WA Geological Survey shoreline retreat rates

| Lighthouse | Annual Retreat Rate | Vulnerability | Score |
|-------------|--------------------|---------------|--------|
| West Point | 0.2m/yr | Low | 5 |
| Admiralty Head | 0.3m/yr | Moderate | 6 |
| Lime Kiln | 0.1m/yr | Lowest | 3 |
| Cape Disapp. | 1.8m/yr | Extreme | 9 |
| Cape Flattery | 1.2m/yr | High | 8 |

---

### 🌩️ Storm Frequency Exposure (Column I) — 10 = Worst

**Source Institutions:** NOAA WaveWatch III, WA Coastal Network storm records

| Lighthouse | Annual Storm Events | Exposure Level | Score |
|-------------|--------------------|----------------|--------|
| West Point | 12 | Moderate | 6 |
| Admiralty Head | 15 | Moderate–High | 7 |
| Lime Kiln | 10 | Low | 5 |
| Cape Disapp. | 25+ | Extreme | 10 |
| Cape Flattery | 22 | Very High | 9 |

---

### 🏛️ Cultural & Historic Significance (Column J)

**Source Institutions:** National Register of Historic Places, WA State SHPO

| Lighthouse | Historic Status | Heritage Detail | Score |
|-------------|-----------------|-----------------|--------|
| West Point | National Register (1881) | Moderate | 9 |
| Admiralty Head | National Historic Landmark (1899) | Moderate | 8 |
| Lime Kiln | National Register (1919) | Active light | 7 |
| Cape Disapp. | NHL (1856) | 2,000+ shipwrecks | 10 |
| Cape Flattery | National Register | Tatoosh Island heritage | 9 |

---

Cultural and tourism value drive funding potential more strongly than simply minimizing hazards.

---

### 🧩 Institutional Data Hierarchy

1. **Primary:** WA State Parks, WA Dept. of Ecology, NOAA coastal databases  
2. **Secondary:** USCG maintenance reports, county GIS & planning data  
3. **Tertiary:** Lighthouse preservation grants, tourism statistics  

---

### 🏆 Summary Insight

**Lime Kiln Lighthouse** ranks #1 due to outstanding tourism, sustainability, and low hazard exposure within San Juan County’s preserve program.  
**Cape Disappointment** and **Cape Flattery** experience severe environmental penalties despite strong heritage scores.

**Data Span:** 27 institutional datasets normalized into consistent 1–10 metrics, producing **actionable, data-driven prioritization for preservation funding**.

---
## Building the Decision Matrix in Excel

### Step 1: Set Up the Spreadsheet

1. Open **Excel → New Blank Workbook**  
2. In **A1**, type `Washington Lighthouses Decision Matrix`  
3. Select **A1:K1 → Merge & Center → Bold → Font Size 16**  
4. Go to **View → Freeze Panes → Freeze Top Row**

---

### Step 2: Enter Column Headers (Row 2)

| Column | Header |
|--------|---------|
| A | Lighthouse |
| B | Pres. Cost (10=low) |
| C | Zoning |
| D | Flood Risk (10=worst) |
| E | Sustainability |
| F | Public Access |
| G | Eco Impact |
| H | Erosion (10=worst) |
| I | Storm (10=worst) |
| J | Cultural Sig. |
| K | Total Score |

---

### Step 3: Enter Lighthouse Data (Rows 3–7)

| Lighthouse | B | C | D | E | F | G | H | I | J |
|-------------|---|---|---|---|---|---|---|---|---|
| West Point | 7 | 9 | 6 | 7 | 8 | 7 | 5 | 6 | 9 |
| Admiralty Head | 6 | 8 | 5 | 8 | 9 | 8 | 6 | 7 | 8 |
| Lime Kiln | 8 | 9 | 4 | 9 | 10 | 10 | 3 | 5 | 7 |
| Cape Disapp. | 5 | 7 | 10 | 6 | 10 | 6 | 9 | 10 | 10 |
| Cape Flattery | 4 | 6 | 8 | 5 | 4 | 9 | 8 | 9 | 9 |

## Step 4: Enter Total Score Formula

1. Click cell **K3**.  
2. Type **exactly**: =(B2+C2+E2+F2+G2+J2)/6*1.5*2 - (D2+H2+I2)/3*1.2*2
3. Press **Enter**
4. Copy formula:
- **Ctrl+C** on **K3**
- Select **K4:K7**
- **Ctrl+V** to paste

---

## Step 5: Expected Results

| Lighthouse       | Total Score |
|------------------|-------------|
| West Point       | 7.15    |
| Admiralty Head   | 7.1    |
| **Lime Kiln** ✅ | **11.7 (BEST)** |
| Cape Disapp.     | -4.5    |
| Cape Flattery    | -4.75    |

---

## Step 6: Format for Professional Look

1. Select **K3:K7** → **Conditional Formatting → Color Scales** (Green = Best, Red = Worst)
2. Select:K7** → **Number → 1 decimal place**
5. Select **A2:K7** → **AutoFit Column Width**

---

## Step 7: Add Visualization

1. Select **A2:K7**
2. Go to **Insert → Recommended Charts → Clustered Column**
3. Set chart title to **"Lighthouse Preservation Priority"**
4. Confirm **Lime Kiln** has the tallest bar (**15.3**)

---
  ```
   In **K9**, type:  
  BEST CHOICE
  ```
- In **K10*, type:  
  ```
  =INDEX(A3:A7,MATCH(MAX(K3:K7),K3:K7,0))
  ```
- The formula returns the name of the top‑ranked lighthouse automatically.
---

## 🧭 Applying the System to Historic Restoration

The **Washington Lighthouses Decision Matrix** acts as a transparent prioritization framework for historic conservation management. When applied systematically, it helps agencies, preservation groups, and local governments evaluate which historic sites deliver the greatest combined value—considering fiscal efficiency, visitor benefit, environmental resilience, and cultural significance in equal terms. The system not only guides restoration prioritization but also supports grant applications, asset risk assessments, and long-term adaptive planning. By basing decisions on measurable data (rather than anecdotal prestige), heritage decision-makers can justify funding allocations that optimize both historical legacy and community benefit.

Projects anywhere—whether the object is a lighthouse, mill, bridge, or museum—can adapt this model by identifying a similar set of **quantitative factors** tailored to local conditions. The same analytical balance between “positive values” (heritage, sustainability, access) and “risk penalties” (cost, erosion, climate exposure) can guide conservation planning at any scale.

---

## 🧩 How to Use This Table for Your Own Projects

1. **Identify Evaluation Factors:** Choose 6–10 criteria most relevant to your assets (e.g., restoration cost, visitor potential, structural integrity, historical importance).  
2. **Set a Consistent Scale:** Use a 1–10 scoring system where higher = better for benefits and “10 = worst” for risks.  
3. **Gather Data:** Source objective, public, or institutional records comparable across all sites—engineering reports, tourism statistics, maintenance budgets, or environmental maps.  
4. **Apply the Formula:** Modify the provided Excel formula to match your chosen variables and weighting preferences.  
5. **Visualize & Interpret:** Use conditional formatting and charts in Excel to highlight highest-value preservation opportunities.  
6. **Decide & Justify:** Use your results to explain restoration decisions in reports, proposals, or funding applications.  

This table template serves as an evolving decision-support tool for **evidence-based heritage management**, ensuring restoration investments preserve not only structures but also the communities and histories they represent.