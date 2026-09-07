# E-Commerce Customer Churn & Cohort Analysis
### Advanced Excel Portfolio Project | Power Query · Power Pivot · DAX · Interactive Dashboard

![Dashboard Preview](images/dashboard-preview.png)

---

## 📌 Project Overview

A UK-based online gift-ware retailer wanted to understand **who their most valuable customers are, how fast customers churn, and where revenue is concentrated** — using two years of raw, unfiltered transaction data.

This project takes **1.07 million rows of messy, real-world retail transactions** and turns them into a fully interactive, single-page Excel dashboard — built entirely with native Excel tools, no external add-ins or software.

**Business questions answered:**
- How long do customers stick around before they stop buying?
- Which customers drive the majority of revenue, and which have gone quiet?
- What percentage of the customer base has churned, and how much revenue is at risk?

---

## 🛠️ Tools & Techniques

| Category | Tools Used |
|---|---|
| **Data Cleaning (ETL)** | Power Query — filtering, type correction, custom columns, text cleaning |
| **Data Modeling** | Power Pivot / Data Model (805K+ rows, optimized for performance) |
| **Analysis Engine** | DAX — `CALCULATE`, `ALLEXCEPT`, `FIRSTNONBLANK`, `DATEDIFF`, `DISTINCTCOUNT` |
| **Customer Scoring** | Advanced formulas — `IFS`, `PERCENTILE`, `XLOOKUP`, `LARGE` + `INDEX/MATCH` |
| **Visualization** | PivotTables, PivotCharts, conditional formatting heatmaps |
| **Dashboard Build** | Linked shapes, grouped objects, dual-cache slicers, single-page interactive layout |

---

## 📊 Dataset

**Source:** [Online Retail II (UCI Machine Learning Repository)](https://archive.ics.uci.edu/dataset/502/online+retail+ii) — real transactional data from a UK-based online retailer, December 2009 to December 2011.

| Metric | Value |
|---|---|
| Raw transaction rows | 1,067,371 |
| Cleaned transaction rows | 805,549 (75.5% retained) |
| Unique customers | 5,878 |
| Date range | Dec 2009 – Dec 2011 (2 years) |
| Countries represented | 40+ |

**Cleaning steps applied:** removed cancelled orders (invoices starting with "C"), removed rows with missing Customer ID, removed non-positive Quantity/Price values, corrected data types, engineered a `Sales` field, and standardized text fields.

---

## 🔑 Key Insights

### 1. Revenue is concentrated in a small, loyal core
**Champions represent 21.6% of customers but generate 68% of total revenue** — a clear Pareto pattern. Loyal Customers over-index too, contributing 14.9% of revenue from 19.3% of the customer base.

### 2. Churn is real, but it's concentrated in low-value customers
**50.9% of customers have churned** (no purchase in 90+ days), but churned customers account for only **19.6% of total historical revenue** (£3.47M of £17.74M) — meaning churn is disproportionately hitting infrequent, lower-value buyers rather than core customers.

### 3. Retention drops off sharply after month one
Cohort analysis shows the steepest customer drop-off happens in the **first month** after acquisition — average retention falls from 100% to ~23% almost immediately. From there it declines more gradually, dipping below 15% around month 9–10 and below 10% by month 18, with some month-to-month fluctuation likely tied to seasonal repurchase patterns typical of a gift retailer. The earliest cohort (Dec 2009) — the only one with a full 24-month view — still retained ~20% of customers at month 24.

### 4. A meaningful "win-back" opportunity exists
The **"Lost"** segment (23.3% of customers, fully churned) still represents **2% of historical revenue** — a pool of past value that's gone cold. Combined with "At Risk" customers (12% of customers, 100% churned, 8.8% of revenue), there's a clear, quantifiable case for a targeted win-back campaign.

---

## 📈 Dashboard Preview

The final deliverable is a **single-page interactive dashboard** featuring:
- 4 live KPI cards (Total Revenue, Total Customers, Churn Rate, Avg Order Value)
- A cohort retention heatmap (25 monthly cohorts × 24 months)
- Revenue share by RFM customer segment
- Active vs. Churned breakdown by segment
- Top 10 customers by lifetime revenue
- Dual Country slicers enabling cross-filtering across all visuals

---

## 🧩 Methodology Notes

- **RFM Segmentation**: customers scored 1–5 on Recency, Frequency, and Monetary value using quintile-based thresholds, then classified into six segments (Champions, Loyal Customers, At Risk, Lost, New/Promising, Needs Attention).
- **Churn definition**: any customer with no purchase in the 90 days prior to the dataset's final date is flagged "Churned." This is a right-censored dataset (it ends Dec 2011), so customers active near the cutoff will appear churned regardless of true future behavior — a known limitation of any fixed-window churn analysis.
- **Dual-slicer architecture**: Excel slicers cannot connect across OLAP (Data Model) and non-OLAP (worksheet range) pivot caches. Rather than rebuild the entire analysis inside the Data Model, two synchronized, identically-styled Country slicers were used to filter across both cache types — a practical workaround for a real Excel architectural constraint.

---

## 📁 File Structure

| Sheet | Contents |
|---|---|
| `Dashboard` | Final interactive single-page view |
| `Cohort_Analysis` | Monthly cohort retention matrix + heatmap |
| `RFM_Scores` | Customer-level R/F/M scores, segments, churn flags |
| `RFM_Summary` | Segment-level customer/revenue breakdown |
| `Churn_Summary` | Churn breakdown by segment |
| *(Data Model)* | Cleaned 805K-row transaction table + DAX measures |

---

## 🚀 About This Project

Built end-to-end as a self-directed portfolio project to demonstrate advanced Excel proficiency: ETL with Power Query, data modeling with Power Pivot, DAX-based analysis, and interactive dashboard design — all using tools available in standard Microsoft Excel.

**Connect with me:** [Your LinkedIn] · [Your Email] · [Your Portfolio Site]
