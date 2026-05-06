# Asheville Airbnb Market Analysis
### Excel Data Analysis Portfolio Project

![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)
![Data](https://img.shields.io/badge/Records-318%2C549%20Reviews-blue?style=flat)
![Listings](https://img.shields.io/badge/Listings-2%2C876-orange?style=flat)
![Status](https://img.shields.io/badge/Status-Complete-success?style=flat)

---

## 📌 Project Overview

This project analyzes Airbnb activity in Asheville, NC on behalf of a fictional travel startup evaluating whether to expand its services into the market. Using three raw datasets from [Inside Airbnb](http://insideairbnb.com/), I conducted a full end-to-end analysis covering pricing trends, neighbourhood demand, host behaviour, and seasonal patterns — delivering six strategic recommendations backed by data.

---

## 🎯 Business Question

> *"We're exploring Asheville as a target market. What do pricing trends look like? Which neighbourhoods and hosts are most active? What does review activity say about demand?"*

---

## 📁 Dataset

| File | Records | Description |
|---|---|---|
| `listings.csv` | 2,876 rows | Listing details including price, room type, host info, neighbourhood |
| `reviews.csv` | 318,549 rows | Individual guest reviews with dates |
| `neighbourhoods.csv` | 8 rows | ZIP code reference for Asheville neighbourhoods |

**Source:** Inside Airbnb — Asheville, NC
**Tool used:** Microsoft Excel Online & Microsoft Excel Desktop 
(Pivot Tables, FILTER, SUMPRODUCT, VLOOKUP, conditional formatting, 
PivotCharts)

---

## 🔍 Methodology

### Step 1 — Data Preparation
- Imported all three CSV files into separate Excel sheets
- Documented missing values using `COUNTBLANK` formulas in a dedicated Data Notes sheet
- Flagged price outliers using an `IF` formula — 24 listings above $2,000 excluded from pricing averages
- Merged `reviews.csv` with `listings.csv` using `VLOOKUP` to connect each review to its neighbourhood
- Added `Review_Year` and `Review_Month` helper columns using `YEAR()` and `MONTH()` functions

**Key data quality findings:**

| Column | Issue | Decision |
|---|---|---|
| neighbourhood_group | 100% empty (2,876 blanks) | Excluded entirely |
| license | 100% empty (2,876 blanks) | Excluded entirely |
| price | 340 missing values | Excluded from averages only |
| last_review | 263 missing values | Kept — listings with no reviews yet |

### Step 2 — Descriptive Analysis
Built pivot tables to calculate average and median nightly price, room type distribution, and listing counts by neighbourhood.

### Step 3 — Pricing Insights
Compared average and median price across all 8 neighbourhoods. Added conditional median using the `FILTER` function (Excel Online compatible). Analyzed the effect of minimum nights on average price using a custom grouping column.

### Step 4 — Review & Demand Analysis
Analyzed 318,549 reviews across years and months to identify demand trends, COVID impact, recovery patterns, and seasonal peaks.

### Step 5 — Host Analysis
Identified top hosts by listing count, calculated host concentration metrics using `SUMPRODUCT`, and compared pricing between multi-listing and single-listing hosts.

---

## 📊 Key Findings

### Pricing
- **Average nightly price:** $181 | **Median:** $128
- The $53 gap between average and median indicates luxury outliers inflating the average — the median is a more reliable market benchmark
- **Most expensive neighbourhood:** 28704 Arden/SW Asheville at $224 average
- **Most active neighbourhood:** 28806 West Asheville with 807 listings
- Listings requiring 31+ night minimums charge 40% less per night ($109) than 1-night listings ($181) — hosts discount for long-term renters

### Room Types
- **89%** of all listings are entire homes (2,571 of 2,876)
- Private rooms account for 10% (285 listings)
- Hotel rooms (1%) and shared rooms (0%) have negligible market presence

### Demand & Reviews
- Reviews grew from **6 in 2011** to **54,540 in 2023** — representing 800%+ market growth
- 2020 saw a slight COVID dip followed by the **strongest single-year growth on record in 2021** (+62%)
- **Peak season:** June, July, August — review volume is 40% higher than winter months
- **Slowest months:** January and February
- West Asheville (28806) generates the highest review volume — most guest activity of any neighbourhood

### Hosts
- **1,019 unique hosts** across 2,876 listings
- 438 multi-listing hosts (43%) vs 581 single-listing hosts (57%)
- Top host — **Towns** — controls **108 listings (4% of the entire market)**
- Top 5 hosts combined control **299 listings (10.4% of supply)**
- Multi-listing hosts charge **$42 more per night** than single-listing hosts ($201 vs $159) — a 26% premium

 ## 📸 Screenshots

### Executive Summary
![Summary](01_summary.png)

### Strategic Recommendations
![Recommendations](02_recommendations.png)

### Annual Review Trend
![Line Chart](04_line_chart.png)

### Room Type Distribution
![Pie Chart](03_pie_chart.png)

### Descriptive Analysis Overview
![Analysis](05_analysis_overview.png) 

---

## ✅ Strategic Recommendations

| Priority | Category | Recommendation |
|---|---|---|
| 🟢 HIGH | Neighbourhood Target | Focus on 28806 (West Asheville) and 28801 (Downtown) first — most listings and highest review volume |
| 🟢 HIGH | Property Type | Prioritize entire home listings exclusively — 89% market dominance, highest avg price |
| 🟢 HIGH | Pricing Strategy | Target $159–$201/night based on host type benchmarks |
| 🟡 MEDIUM | Seasonal Strategy | Launch campaigns before June peak — build promotions for Nov–Feb off-season |
| 🔴 RISK | Host Concentration | Monitor dependency on Towns (4%) and top 5 hosts (10.4%) — supply risk if operators exit |
| 🔴 RISK | Seasonal Demand | Build dynamic pricing tools before winter — reviews drop 40% Nov–Feb |

### Overall Verdict
**Asheville is a strong expansion market.** Demand has grown consistently for over a decade, the market is not oversaturated, and multiple neighbourhoods show clear opportunity. Focus on West Asheville and Downtown with entire home listings priced $159–$201/night, targeting the summer peak season (June–August).

---

## 📂 Repository Structure

```
asheville-airbnb-analysis/
│
├── README.md                          ← You are here
│
├── data/
│   ├── listings.csv                   ← Raw listings data (2,876 rows)
│   ├── reviews.csv                    ← Raw review data (318,549 rows)
│   └── neighbourhoods.csv             ← ZIP code reference
│
├── Asheville_Airbnb_Analysis.xlsx     ← Full Excel workbook
│
└── screenshots/
    ├── recommendations.png            ← Strategic recommendations sheet
    ├── annual_review_trend.png        ← Line chart — demand 2011–2025
    ├── room_type_distribution.png     ← Pie chart — room types
    └── analysis_summary.png          ← Pivot table overview
```

---

## 🛠️ Excel Skills Demonstrated

- Pivot Tables with custom sorting, filtering, and grouping
- `FILTER` function for conditional median calculations (Excel Online)
- `SUMPRODUCT` for counting unique and conditional values
- `VLOOKUP` and `IFERROR` for cross-sheet data merging
- `YEAR()`, `MONTH()`, `TEXT()`, `DATE()` for date parsing
- Nested `IF` statements for custom categorization
- `COUNTBLANK` for data quality auditing
- PivotCharts — line chart, pie chart, bar chart
- Conditional formatting and color-coded dashboards
- Data quality documentation and decision logging

---

## 💡 Challenges & Solutions

| Challenge | Solution |
|---|---|
| Excel Online does not support legacy array formulas (`Ctrl+Shift+Enter`) | Replaced with modern `FILTER` function for conditional medians |
| `neighbourhood_group` column 100% empty | Documented and excluded — used `neighbourhood` (ZIP code) instead |
| Pivot Tables cannot calculate median natively | Used `=MEDIAN(FILTER(...))` alongside pivot tables |
| Price outliers skewing averages | Flagged with `Price_Flag` column — excluded listings above $2,000 from averages |

---

## 👤 About

**Role:** Data Analyst (Portfolio Project)
**Context:** Travel research firm evaluating Asheville, NC as an expansion market for a travel startup
**Data source:** [Inside Airbnb](http://insideairbnb.com/) — publicly available Airbnb activity data

---

*This project was completed as part of a data analyst portfolio. 
Analysis and modelling was performed in Microsoft Excel Online. 
Final design, formatting, and visualizations were produced in 
Microsoft Excel Desktop.*
