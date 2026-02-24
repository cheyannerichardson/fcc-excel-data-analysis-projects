## Complete Setup: Stanley Seawall Traffic Inventory Project

Transform a basic car inventory Excel template into a professional urban planning traffic survey tool for Vancouver's Stanley Park Seawall.

## Phase 1: Field Planning (Day Before Survey)
### 1. Select Survey Points
```
Hotspots (4 locations, 15-min intervals):
- Pipeline Road (vehicle pinchpoint)
- Second Beach (pedestrian heavy) 
- Ferguson Point (cyclist-ped conflicts)
- Causeway (all modes converge)
```

### 2. Create Field Data Sheet
**File → New → Blank Workbook → Save As "SeawallTraffic_YYYYMMDD.xlsx"**
```
Sheet1 "Seawall_Data" headers (A1:H1):
ID | Type | Location | Count | Time | Direction | Status | Notes
```
**Pre-fill dropdowns** (Data → Data Validation):
- **Type** (B:B): `Vehicle,Bike,Ped`
- **Location** (C:C): `Pipeline Rd,Second Beach,Causeway,Ferguson Pt`
- **Direction** (F:F): `CCW,CW`
- **Status** (G:G): `Normal,Slow,Conflict,Stopped`

### 3. Format for Field Use
```
- Freeze top row: View → Freeze Panes → Freeze Top Row
- Bold headers, yellow fill (Home → Fill Color)
- Auto-fit columns
- Add 100 blank rows
```

## Phase 2: Field Data Collection (Survey Day)
### 4. Equipment Setup (2-person team per location)
```
Person 1: Tally marks on paper → enter every 15 mins
Person 2: Phone timer + photos of conflicts
Timing: Noon-2PM Saturday (peak tourist hours)
Interval: Count 15 mins, enter row, repeat
```

### 5. Counting Protocol
```
Each row = ONE 15-minute batch:
- Type: Vehicle=car/taxi/bus, Bike=all cycles, Ped=walkers/joggers
- Count: Total passing in 15 mins (direction specific)
- Status: Conflict=near-miss/block, Slow=backup, Normal=steady
- Notes: "Tour group 8ppl", "Rental bikes wrong way"
```

## Phase 3: Excel Analysis Setup (Post-Survey)
### 6. Convert to Table
```
Select A1:H100 → Insert → Table → check "Headers" → OK
Table Design → Table Name: TrafficData
```

### 7. Add Visual Alerts
**Column G (Status)**:
```
Select G2:G100 → Home → Conditional Formatting → New Rule
→ Use formula: =$G2="Conflict" → Format → Yellow fill → OK
```

### 8. Create 3 Analysis Sheets
**Sheet2 "Flow_Analysis"**:
```
Insert → PivotTable → New Worksheet
Time → Rows (right-click → Group → Hours only → Format: h"PM")
Type → Columns  
Count → Values (Sum)
Location → Filters
Insert → Slicer → Type (Bike button interactivity)
```

**Sheet3 "Conflicts_Analysis"**:
```
Location → Rows
Type → Columns
Status → Filters (Conflict only)
Count → Values (Sum)
Insert → PivotChart → Clustered Column
```

**Sheet4 "Dashboard"**:
```
J1: =SUMIF(TrafficData[Status],"Conflict",TrafficData[Count])
K1: =SUM(TrafficData[Count])
L1: =J1/K1 (conflict %) 
```

## Phase 4: Generate Deliverables
### 9. Key Reports (Copy-paste values)
```
Dashboard → Filter Second Beach → Screenshot
Conflicts chart → Ferguson Pt → Save as PNG
Flow pivot → Noon peak → Export table
```

### 10. Present to Park Board
```
1-slide summary:
"Peak conflicts: 36 total, 58% Ferguson Pt cyclist-ped drifts"
"Recommendation: Bollards + signage at 3 hotspots"
"Cost: $8K materials, 2-day install"
```

## Pro Tips for Vancouver Seawall
```
- Survey May-Sep (tourist season)
- CCW = 85% cyclist flow (standard)
- Conflicts peak weekends 12-2PM
- Rental bikes = primary wrong-way offenders
- Export to Google Earth for location visuals
```

**Total setup time**: 45 mins. **Field time**: 4 locations × 2hrs = 8 person-hours. **Analysis**: 30 mins. **Ready for stakeholder meeting**.
