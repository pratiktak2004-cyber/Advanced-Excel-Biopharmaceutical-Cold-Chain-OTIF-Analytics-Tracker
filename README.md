# Advanced Excel Biopharmaceutical Cold-Chain & OTIF Analytics Tracker

An enterprise-grade supply chain analytics workbook designed in Microsoft Excel to model, monitor, and evaluate biopharmaceutical shipments, thermal compliance, and On-Time In-Full (OTIF) delivery SLAs.

---

## Executive Dashboard & Analytics Preview

> *Place your dashboard screenshot (`dashboard_preview.png`) in the root directory to display here.*

![Dashboard Preview](dashboard_preview.png)

---

## Executive Summary & Project Goal
In biopharmaceutical supply chains, keeping biotherapeutics (such as monoclonal antibodies, vaccines, and cell therapies) within strict thermal windows during transit is critical. Temperature excursions risk batch degradation, compliance violations, and severe financial losses.

This project implements a normalized relational data model in **Advanced Excel** to automate cross-table data lookups, flag cold-chain breaches in real time via formula-driven alerts, monitor delivery performance metrics, and deliver executive-level analytics.

---

## Workbook Architecture & Sheet Mapping

The workbook is structured into a clean relational database schema separating master dimension datasets from operational transaction logs:

| Sheet Name | Original Tab | Role & Contents |
| :--- | :--- | :--- |
| **`Dashboard`** | Sheet4 | Executive control panel containing top-level KPI metric cards and interactive lookup controls. |
| **`Analytics`** | Analytics | Dedicated reporting view housing Pivot Tables and visual PivotCharts. |
| **`Fact_Shipments`** | Sheet6 | Primary transaction log containing shipment IDs, dates, `XLOOKUP` formulas, and alerts. |
| **`Dim_Biotherapeutic`** | Sheet10 | Master reference dimension for product specifications, classes, and thermal limits. |
| **`Dim_Distribution_Hub`** | Sheet5 | Master reference dimension for logistics hubs, regions, and temperature corridors. |

---

## Advanced Excel Technical Stack & Formulas

### 1. Executive Dashboard & Dynamic Lookups (`Dashboard` Sheet)
* **Product Price Dynamic Lookup (`B3`):**
  ```excel
  =XLOOKUP(B1, Dim_Biotherapeutic[product_id], Dim_Biotherapeutic[unit_price], "Not Found")
