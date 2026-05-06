# Production Operations Dataset — Distillery Analytics Portfolio

**Author:** Tarun Dev Dinesh  
**Tool:** Microsoft Excel (Advanced)  
**Domain:** Regulated Industrial Operations — Spirits Production  
**Dataset Size:** 670 rows across 11 sheets (5 fact tables, 3 dimension tables, 2 dashboards, 1 joined view)  
**Period Covered:** 2023 (full calendar year)  
**Purpose:** Portfolio project demonstrating applied Excel analytics skills in an operational data context

---

## Overview

This dataset simulates a multi-site spirits distillery operation, covering the full production workflow from raw material intake through to quality control, equipment maintenance, and operator performance. It is structured as a relational data model — mirroring how operational data is managed in ERP environments such as SAP HANA — and is designed to support a range of Excel analytical exercises including Pivot Tables, Power Query, DAX measures, XLOOKUP, variance analysis, and dashboard development.

The data reflects realistic operational patterns: yield variances, QC failures, downtime events, raw material consumption tracking, maintenance scheduling, and multi-site production comparisons. All site, operator, and batch references are fictional.

---

## File Structure

### Dimension Tables (Reference / Lookup Data)

| Sheet | Description | Key Fields |
|---|---|---|
| `Sites` | Master list of 5 production sites with region, function, and daily capacity | `site_id`, `site_name`, `region`, `capacity_L_per_day` |
| `Operators` | Operator roster with role, primary site, and years of experience | `operator_id`, `operator_name`, `role`, `primary_site` |
| `Equipment` | Equipment register across sites with asset type and status | `equipment_id`, `site_id`, `equipment_type`, `status` |

### Fact Tables (Transactional / Operational Data)

| Sheet | Description | Key Fields |
|---|---|---|
| `Batch_Production` | Core production log — one row per batch run across all sites and spirit types | `batch_id`, `site_id`, `spirit_type`, `unit_operation`, `operator_id`, `scheduled_output_L`, `actual_output_L`, `variance_L`, `variance_pct`, `downtime_hrs`, `QC_Final`, `erp_batch_ref` |
| `QC_Instrument_Readings` | Lab QC readings per batch — deviation from target across parameters such as ABV, copper, pH | `reading_id`, `batch_ref`, `qc_parameter`, `target_value`, `measured_value`, `deviation_pct`, `within_spec` |
| `Raw_Material_Consumption` | Material intake and consumption per batch — grain, water, yeast, and utilities | `record_id`, `batch_ref`, `material_name`, `qty_ordered`, `qty_consumed`, `variance_qty`, `unit_cost_CAD`, `total_cost_CAD` |
| `Equipment_Maintenance` | Work orders for scheduled and corrective maintenance across sites | `work_order_id`, `site_id`, `equipment_id`, `maintenance_type`, `priority`, `duration_hrs`, `labour_cost_CAD`, `total_cost_CAD`, `downtime_caused_hrs`, `status` |
| `Anomaly_Incident_Log` | Flagged anomalies and incidents — batch holds, ABV deviations, yield exceptions | `incident_id`, `batch_ref`, `incident_type`, `severity`, `resolution_status` |

### Joined View

| Sheet | Description |
|---|---|
| `Production_With_Site_Location` | Batch_Production joined to Sites — flattened view for reporting without XLOOKUP dependency |

### Dashboards

| Sheet | Description |
|---|---|
| `Production Dashboard` | Pivot Table summary — monthly output vs scheduled, downtime by site, QC pass rates by spirit type, energy efficiency |
| `Operator Dashboard` | Operator-level performance scorecard — batch count, average yield variance, QC pass rate, downtime hours |

---

## Data Model Relationships

```
Sites ──────────────┬──── Batch_Production ────┬──── QC_Instrument_Readings
                    │           │               └──── Raw_Material_Consumption
Operators ──────────┘           │               └──── Anomaly_Incident_Log
                                │
Equipment ──────────────────── Equipment_Maintenance
```

**Primary keys:** `site_id`, `operator_id`, `equipment_id`, `batch_id`  
**Foreign keys:** `batch_ref` links QC, materials, and anomaly records back to Batch_Production

---

## Excel Skills Demonstrated

### Lookup and Reference
- `XLOOKUP` — used to pull site names, operator names, and region into the joined production view from dimension tables
- `VLOOKUP` — applied for equipment cross-referencing in maintenance records
- Named ranges and structured table references (`Table[Column]`) used throughout for formula readability and reliability

### Pivot Tables and Pivot Charts
- Monthly output summary (scheduled vs actual litres) with Grand Total row
- Downtime hours by site and by downtime reason — Count and Sum pivots
- QC Final Pass % by spirit type
- Energy efficiency by site and spirit type
- Operator scorecard: batch count, average variance, QC pass rate (Operator Dashboard sheet)

### Power Query
- Data loaded and transformed via Power Query connections
- Column types enforced (date, numeric, text) at load step
- `variance_pct` calculated as a custom column: `= [variance_L] / [scheduled_output_L]`
- Downtime reason null values replaced with "No Downtime" for cleaner pivot behaviour

### DAX Measures (for Power Pivot / Power BI extension)
- `QC Pass Rate = DIVIDE(COUNTIF([QC_Final],1), COUNTA([QC_Final]))`
- `Yield Variance % = AVERAGE([variance_pct])`
- `Energy Intensity = DIVIDE(SUM([energy_consumed_GJ]), SUM([actual_output_L]))`

### Variance Analysis
- Batch-level yield variance (litres and %) — positive and negative deviations flagged
- ABV target vs actual tracked per batch via QC_Instrument_Readings
- Raw material consumption variance: ordered vs consumed quantity and cost

### Conditional Formatting
- QC status column: green (PASS) / red (FAIL) traffic light formatting
- Variance % column: colour scale from negative (red) to positive (green) to highlight yield outliers
- Downtime hours: data bars applied to operator scorecard for visual comparison

### Data Validation and QA/QC
- Dropdown validation applied to `spirit_type`, `unit_operation`, `shift`, `QC_Final`, and `status` columns to prevent free-text entry errors
- `qc_status` field cross-validated against `QC_Final` flag for consistency
- `erp_batch_ref` formatted consistently (`ERP-YYYY-XXXX`) and checked for duplicates using conditional formatting

---

## Suggested Analysis Exercises

The following exercises are designed to demonstrate analytical depth using this dataset in Excel, Power Query, or Power BI.

### Beginner
1. **Monthly Production Summary** — Pivot Table of actual output by month and spirit type. Add a slicer for site.
2. **QC Pass Rate by Spirit** — `COUNTIF` / `AVERAGEIF` to calculate pass rate per spirit type without a pivot.
3. **Operator Batch Count** — How many batches did each operator run? Which operator had the highest average yield variance?

### Intermediate
4. **Downtime Root Cause Analysis** — Which downtime reason accounts for the most lost hours? Which site is most affected?
5. **XLOOKUP Join** — Use `XLOOKUP` to bring `site_name` and `region` from the Sites table into Batch_Production without using the pre-joined sheet.
6. **Yield Threshold Flag** — Add a helper column flagging batches where `variance_pct` exceeds ±5%. What proportion of batches breach this threshold?
7. **Cost per Litre** — Join Raw_Material_Consumption to Batch_Production to calculate total raw material cost per litre of actual output by batch.

### Advanced
8. **Operator Scorecard** — Build a dynamic scorecard using Pivot Tables showing: batch count, QC pass rate, average yield variance %, and total downtime hours per operator. Add conditional formatting.
9. **Energy Intensity Trend** — Calculate `energy_consumed_GJ / actual_output_L` per batch. Trend this by month using a line chart. Identify the most and least energy-efficient spirit types.
10. **Maintenance Cost vs Downtime Correlation** — Join Equipment_Maintenance to Batch_Production via `site_id`. Is there a relationship between maintenance spend and production downtime? Use a scatter chart.
11. **Power Query: Multi-Table Model** — Load Batch_Production, Sites, and Operators into Power Query. Merge on `site_id` and `operator_id`. Build a single flat table ready for pivot analysis without using XLOOKUP formulas.
12. **Anomaly Deep Dive** — Join Anomaly_Incident_Log to Batch_Production. What types of incidents most frequently coincide with QC failures? Which operators or sites appear most in the log?

---

## Notes on Data Design

- All personal names (operators) are fictional.
- Site names and locations are fictional but geographically plausible for a Canadian multi-site operation.
- Batch dates are stored as Excel serial numbers — format as `DD-MMM-YYYY` for display.
- `energy_consumed_GJ` values are realistic for small-to-mid scale distillery operations.
- QC readings include deliberate out-of-spec values to support anomaly detection exercises.
- Some batches intentionally have missing `downtime_reason` values (blank = no downtime event) — useful for Power Query null handling practice.

---

## Potential Extensions

- **Load into Power BI** — The relational structure maps directly to a Power BI data model. Connect tables via `site_id`, `operator_id`, and `batch_id`/`batch_ref` and recreate the dashboards with interactive slicers.
- **SQL Practice** — Export sheets as CSVs and load into SQLite or Mode Analytics. The fact/dimension structure supports `JOIN`, `GROUP BY`, `HAVING`, window functions (`RANK`, `LAG`), and CTE exercises.
- **Swap with Real Data** — The schema is generic enough to replace simulated values with real operational data from any batch-process manufacturing environment.

---

## Contact

**Tarun Dev Dinesh**  
Calgary, AB  
tdd.nero@gmail.com  
[LinkedIn](https://www.linkedin.com/in/thedatadistiller)

