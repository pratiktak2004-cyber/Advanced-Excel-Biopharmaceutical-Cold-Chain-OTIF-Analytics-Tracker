# Advanced-Excel-Biopharmaceutical-Cold-Chain-OTIF-Analytics-Tracker
Advanced Excel logistics analytics model to, track, and analyze cold-chain biopharmaceutical shipments. The model leverages relational schema design, dynamic array lookups, automated conditional logic, aggregate KPI cards, and PivotTable/PivotChart visual analytics to evaluate temperature compliance and delivery SLAs across distribution corridors.

---

## Executive Dashboard & Analytics Preview

> *Place your dashboard screenshot (`dashboard_preview.png`) in the root directory to display here.*

![Dashboard Preview](dashboard_preview.png)

---

## Executive Summary & Project Goal
In biopharmaceutical supply chains, keeping biotherapeutics (such as monoclonal antibodies, vaccines, and cell therapies) within strict thermal windows during transit is critical. Temperature excursions risk batch degradation, compliance violations, and severe financial losses.

This project implements a normalized relational data model in **Advanced Excel** to automate cross-table data lookups, flag cold-chain breaches in real time via formula-driven alerts, monitor delivery performance metrics, and deliver executive-level analytics.

---

## Technical Architecture & Advanced Excel Features

### 1. Relational Data Modeling & Multi-Table Schema
* Engineered a normalized relational schema separating master dimension datasets from operational logs using native Excel Tables (`Ctrl + T`) for dynamic range expansion and clean structural references.
  * **`Dashboard`** *(Sheet4)*: Executive overview featuring top-line KPI metric cards.
  * **`Analytics`** *(Sheet8)*: Interactive reporting tab driven by Pivot Tables and PivotCharts.
  * **`Fact_Shipments`** *(Sheet6)*: Core operational log recording transit metrics, dispatch dates, and thermal integrity.
  * **`Dim_Biotherapeutic`** *(Sheet2)*: Master product dataset storing unit prices and thermal threshold boundaries.
  * **`Dim_Distribution_Hub`** *(Sheet3)*: Master hub dataset mapping distribution centers and geographic regions.

### 2. Dynamic Array Lookups & Data Integration (`XLOOKUP`)
* Implemented nested `XLOOKUP` functions across relational tables to dynamically fetch unit prices, storage temperature constraints, and distribution hub details into transactional records without hardcoding data.

### 3. Automated Conditional Formatting & Alert Logic
* Formatted rule-based Conditional Formatting to dynamically triage operational risk across dataset rows:
  * **Thermal Excursion Alerts:** Highlighted in Red for out-of-range thermal readings vs. Green for compliant batches.
  * **Delivery SLA Status:** Visual indicator flags for **Delayed** shipments vs. **Delivered-OTIF**.

### 4. Executive KPI Metric Summary Cards
* Designed high-level executive cards using array and aggregate functions (`COUNTA`, `COUNTIF`) with custom cell formatting to compute real-time operational indicators:
  * **Total Shipments Monitored**
  * **OTIF Compliance Rate (%)**
  * **Thermal Compliance Rate (%)**

### 5. Pivot Analytics & Interactive Visualization
* Built a dedicated `Analytics` dashboard utilizing PivotTables and custom PivotCharts to aggregate status distributions, quantify delay frequencies, and visualize logistics performance across distribution corridors.

---

## Repository Workbook Structure

| Sheet Name | Original Tab | Purpose & Contents |
| :--- | :--- | :--- |
| **`Dashboard`** | Sheet4 | Executive summary view containing KPI summary cards and key metric counters. |
| **`Analytics`** | Sheet8 | Reporting dashboard housing Pivot Tables and visual PivotCharts. |
| **`Fact_Shipments`** | Sheet6 | Primary transaction log containing shipment IDs, dates, `XLOOKUP` formulas, and alerts. |
| **`Dim_Biotherapeutic`** | Sheet2 | Master reference dimension for product specifications and thermal limits. |
| **`Dim_Distribution_Hub`** | Sheet3 | Master reference dimension for logistics hubs and regional routes. |

---

## 🚀 How to Run & Explore
1. Clone or download this repository.
2. Open `Biopharma_ColdChain_Tracker.xlsx` in **Microsoft Excel 2019, 365, or later** (for native `XLOOKUP` support).
3. Navigate to **`Dashboard`** for top-level operational metrics or **`Analytics`** to interact with the PivotTable and PivotChart.
