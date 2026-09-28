# Healthcare Supply Chain Inventory & Reconciliation System

A simulated healthcare supply-chain inventory and reconciliation system built in **Excel** to demonstrate practical skills in inventory tracking, receiving, usage monitoring, physical counts, reconciliation, Min/Max inventory controls, replenishment planning, and quality assurance.

> **Important:** This is a simulated portfolio project using fictional healthcare supply data. It does not use confidential data from any hospital, healthcare organization, or employer.

---

## Project Overview

Healthcare supply departments need accurate inventory records to ensure that supplies are available when needed while maintaining appropriate stock levels.

This project models a simplified point-of-use healthcare inventory workflow:

```text
Receiving
   ↓
Inventory Master
   ↓
Usage / Issuance
   ↓
System Inventory
   ↓
Physical Count
   ↓
Reconciliation
   ↓
Min/Max Monitoring
   ↓
Reorder / Replenishment
   ↓
Quality Assurance
```

The system is designed around common operational tasks such as:

* Recording received supplies
* Tracking supply usage
* Maintaining inventory levels
* Comparing system quantities with physical counts
* Identifying inventory discrepancies
* Monitoring Min/Max levels
* Calculating replenishment quantities
* Performing barcode and data-quality checks
* Supporting accurate operational reporting

---

## Tools

* **Microsoft Excel**
* Excel formulas
* Data validation / dropdown controls
* Conditional formatting
* VLOOKUP
* SUMIF
* IF / IFERROR
* Basic inventory calculations
* Spreadsheet-based data quality checks

No actual hospital inventory system or confidential healthcare data is used.

---

# Workbook Structure

The workbook contains the following tabs:

| Tab                | Purpose                                                                        |
| ------------------ | ------------------------------------------------------------------------------ |
| `Inventory_Master` | Central reference table containing supply information and inventory parameters |
| `Receiving_Log`    | Records incoming supplies and receiving QA information                         |
| `Usage_Log`        | Records supplies issued or consumed                                            |
| `Physical_Counts`  | Records physical inventory counts                                              |
| `Reconciliation`   | Compares system inventory with physical inventory                              |
| `Reorder_Report`   | Monitors Min/Max levels and replenishment requirements                         |
| `QA_Checks`        | Planned data-quality and inventory QA checks                                   |
| `Dashboard`        | Planned summary of inventory and operational KPIs                              |

Steps 1–5 establish the underlying inventory workflow. Step 6 begins the replenishment and Min/Max monitoring workflow.

---

# Step 1 — Inventory Master

The `Inventory_Master` tab serves as the central reference table for the inventory system.

## Columns

| Column | Field         |
| ------ | ------------- |
| A      | `Supply_ID`   |
| B      | `Supply_Name` |
| C      | `Category`    |
| D      | `Location`    |
| E      | `Unit`        |
| F      | `Min_Level`   |
| G      | `Max_Level`   |
| H      | `Current_Qty` |
| I      | `Unit_Cost`   |
| J      | `Barcode`     |
| K      | `Supplier`    |
| L      | `Status`      |

## Simulated Inventory

The project uses fictional supplies representing common healthcare inventory categories.

Examples include:

* IV Start Kits
* Gauze
* Nitrile Gloves
* Surgical Masks
* Alcohol Prep Pads
* Syringes
* IV Extension Sets
* Adhesive Tape

Each item has a simulated:

* Minimum stock level
* Maximum stock level
* Current quantity
* Unit cost
* Barcode
* Supplier
* Storage location

## Inventory Status

The `Status` column identifies whether an item has fallen below its minimum level.

Excel formula:

```excel
=IF(H2<F2,"REORDER","OK")
```

This produces:

* `REORDER` — current quantity is below the minimum
* `OK` — current quantity is at or above the minimum

---

# Step 2 — Receiving Log

The `Receiving_Log` tab records incoming inventory.

## Columns

| Column | Field               |
| ------ | ------------------- |
| A      | `Receipt_ID`        |
| B      | `Receipt_Date`      |
| C      | `Supply_ID`         |
| D      | `Supply_Name`       |
| E      | `Location`          |
| F      | `Quantity_Received` |
| G      | `Unit`              |
| H      | `Supplier`          |
| I      | `Purchase_Order`    |
| J      | `Barcode`           |
| K      | `Received_By`       |
| L      | `QA_Status`         |
| M      | `Notes`             |

## Workflow

A receiving transaction contains:

1. Receipt ID
2. Date received
3. Supply ID
4. Quantity received
5. Supplier
6. Purchase order
7. Barcode
8. Person receiving the shipment
9. QA status
10. Notes

Reference information such as supply name, location, unit, supplier, and barcode is automatically retrieved from `Inventory_Master`.

### Example Lookup Formulas

Supply name:

```excel
=IFERROR(VLOOKUP(C2,Inventory_Master!A:K,2,FALSE),"")
```

Location:

```excel
=IFERROR(VLOOKUP(C2,Inventory_Master!A:K,4,FALSE),"")
```

Unit:

```excel
=IFERROR(VLOOKUP(C2,Inventory_Master!A:K,5,FALSE),"")
```

Supplier:

```excel
=IFERROR(VLOOKUP(C2,Inventory_Master!A:K,11,FALSE),"")
```

Barcode:

```excel
=IFERROR(VLOOKUP(C2,Inventory_Master!A:K,10,FALSE),"")
```

## QA Status

The receiving QA dropdown contains:

* `Passed`
* `Exception`
* `Pending`

Conditional formatting highlights exceptions and pending records for review.

---

# Step 3 — Usage Log

The `Usage_Log` records inventory issued or consumed during operations.

## Columns

| Column | Field           |
| ------ | --------------- |
| A      | `Usage_ID`      |
| B      | `Usage_Date`    |
| C      | `Supply_ID`     |
| D      | `Quantity_Used` |
| E      | `Destination`   |
| F      | `Issued_By`     |
| G      | `QA_Status`     |
| H      | `Notes`         |

The transaction log intentionally contains only the information necessary to record an issuance/usage event.

Reference information is retrieved later by the reconciliation and reporting tabs rather than duplicating it throughout the transaction log.

## Controls

### Supply ID

`Supply_ID` uses a dropdown populated from:

```text
Inventory_Master!A2:A11
```

### Quantity

`Quantity_Used` must be greater than zero.

Custom validation:

```excel
=D2>0
```

Invalid entries should be rejected.

### QA Status

Dropdown values:

* `Passed`
* `Pending`
* `Exception`

---

# Step 4 — Reconciliation

The `Reconciliation` tab calculates the expected system inventory and compares it with physical inventory.

## Columns

| Column | Field                   |
| ------ | ----------------------- |
| A      | `Supply_ID`             |
| B      | `Supply_Name`           |
| C      | `Location`              |
| D      | `Min_Level`             |
| E      | `Max_Level`             |
| F      | `Beginning_Qty`         |
| G      | `Total_Received`        |
| H      | `Total_Used`            |
| I      | `System_Qty`            |
| J      | `Physical_Count`        |
| K      | `Variance`              |
| L      | `Variance_%`            |
| M      | `Reconciliation_Status` |

## System Quantity

The core inventory calculation is:

```text
System Quantity =
Beginning Quantity
+ Total Received
- Total Used
```

### Formulas

Supply ID:

```excel
=Inventory_Master!A2
```

Supply name:

```excel
=IFERROR(VLOOKUP(A2,Inventory_Master!A:K,2,FALSE),"")
```

Location:

```excel
=IFERROR(VLOOKUP(A2,Inventory_Master!A:K,4,FALSE),"")
```

Minimum level:

```excel
=IFERROR(VLOOKUP(A2,Inventory_Master!A:K,6,FALSE),"")
```

Maximum level:

```excel
=IFERROR(VLOOKUP(A2,Inventory_Master!A:K,7,FALSE),"")
```

Beginning quantity:

```excel
=IFERROR(VLOOKUP(A2,Inventory_Master!A:K,8,FALSE),"")
```

Total received:

```excel
=SUMIF(Receiving_Log!C:C,A2,Receiving_Log!F:F)
```

Total used:

```excel
=SUMIF(Usage_Log!C:C,A2,Usage_Log!D:D)
```

System quantity:

```excel
=F2+G2-H2
```

---

## Physical Count

The `Physical_Count` field initially remains blank until a physical inventory count is performed.

Once physical counts are entered, the spreadsheet calculates the difference between the recorded system quantity and the observed physical quantity.

### Variance

```excel
=IF(J2="","",J2-I2)
```

Interpretation:

```text
Positive variance = physical count exceeds system quantity
Negative variance = physical count is below system quantity
Zero variance     = physical count matches system quantity
```

### Variance Percentage

```excel
=IF(OR(J2="",I2=0),"",K2/I2)
```

The column is formatted as a percentage.

---

## Reconciliation Status

The project uses a simulated 5% discrepancy threshold for demonstration purposes.

```excel
=IF(J2="","COUNT REQUIRED",
 IF(K2=0,"RECONCILED",
 IF(ABS(L2)>=0.05,"RECOUNT REQUIRED","INVESTIGATE")))
```

Possible results:

* `COUNT REQUIRED`
* `RECONCILED`
* `INVESTIGATE`
* `RECOUNT REQUIRED`

> The 5% threshold is a portfolio-project rule and is **not** intended to represent an actual hospital inventory policy.

---

# Step 5 — Physical Counts

The `Physical_Counts` tab represents the human inventory-counting process.

## Columns

| Column | Field                 |
| ------ | --------------------- |
| A      | `Count_ID`            |
| B      | `Count_Date`          |
| C      | `Supply_ID`           |
| D      | `Location`            |
| E      | `System_Qty_At_Count` |
| F      | `Physical_Count`      |
| G      | `Counted_By`          |
| H      | `Count_Status`        |
| I      | `Notes`               |

## Workflow

A physical inventory count records:

1. Which supply was counted
2. Where it was counted
3. What the system expected
4. What was physically observed
5. Who performed the count
6. Whether the count was complete
7. Any notes or discrepancies

### Supply ID

Dropdown sourced from:

```text
Inventory_Master!A2:A11
```

### Location

Automatically retrieved:

```excel
=IFERROR(VLOOKUP(C2,Inventory_Master!A:K,4,FALSE),"")
```

### System Quantity at Count

Automatically retrieved:

```excel
=IFERROR(VLOOKUP(C2,Reconciliation!A:I,9,FALSE),"")
```

### Physical Count

`Physical_Count` is intentionally a **manual input**.

This represents the actual human counting process and prevents the physical count from simply reproducing the system quantity.

### Count Status

Dropdown:

* `Complete`
* `Pending`
* `Recount Required`

---

## Physical Count → Reconciliation

The `Reconciliation` tab retrieves the latest simulated physical count:

```excel
=IFERROR(VLOOKUP(A2,Physical_Counts!C:F,4,FALSE),"")
```

This allows the reconciliation process to compare:

```text
System Quantity
       ↓
Physical Count
       ↓
Variance
       ↓
Variance %
       ↓
Reconciliation Status
```

## Current Simulated Results

The current dataset produces the following examples:

| Supply | System Qty | Physical Count | Variance | Variance % | Status      |
| ------ | ---------: | -------------: | -------: | ---------: | ----------- |
| S001   |         88 |             85 |       -3 |     -3.41% | INVESTIGATE |
| S002   |        240 |            240 |        0 |      0.00% | RECONCILED  |
| S003   |        147 |            143 |       -4 |     -2.72% | INVESTIGATE |
| S004   |        150 |            150 |        0 |      0.00% | RECONCILED  |
| S005   |        154 |            153 |       -1 |     -0.65% | INVESTIGATE |
| S006   |        112 |            112 |        0 |      0.00% | RECONCILED  |
| S007   |        175 |            171 |       -4 |     -2.29% | INVESTIGATE |
| S008   |        207 |            207 |        0 |      0.00% | RECONCILED  |
| S009   |         84 |             82 |       -2 |     -2.38% | INVESTIGATE |
| S010   |        114 |            114 |        0 |      0.00% | RECONCILED  |

No current discrepancy exceeds the simulated 5% recount threshold.

---

# Step 6 — Min/Max Reorder Report

The `Reorder_Report` begins the replenishment-planning portion of the project.

The goal is to demonstrate how inventory data can be used to identify supplies that require replenishment and determine how much inventory is needed to return the item to its maximum level.

## Planned Columns

| Column | Field                    |
| ------ | ------------------------ |
| A      | `Supply_ID`              |
| B      | `Supply_Name`            |
| C      | `Location`               |
| D      | `Unit`                   |
| E      | `Min_Level`              |
| F      | `Max_Level`              |
| G      | `System_Qty`             |
| H      | `Quantity_to_Max`        |
| I      | `Reorder_Status`         |
| J      | `Unit_Cost`              |
| K      | `Estimated_Reorder_Cost` |

## System Quantity

The report should use the reconciled system quantity rather than relying on a manually entered inventory number.

This keeps the workflow connected:

```text
Receiving + Usage
        ↓
Reconciliation
        ↓
System Quantity
        ↓
Reorder Report
```

## Quantity to Max

The amount needed to restore inventory to the maximum level can be calculated as:

```text
Quantity to Max = Max Level - System Quantity
```

A negative result should not create a negative order quantity, so the report should use zero when the item is already at or above its maximum.

Example:

```excel
=MAX(0,F2-G2)
```

where:

* `F2` = Max Level
* `G2` = System Quantity

## Reorder Status

The report can classify supplies using the project's Min/Max rules:

```text
System Quantity < Min Level → REORDER
System Quantity ≥ Min Level → OK
```

Example:

```excel
=IF(G2<E2,"REORDER","OK")
```

where:

* `G2` = System Quantity
* `E2` = Min Level

## Estimated Reorder Cost

The estimated cost of replenishing to the maximum level can be calculated as:

```text
Quantity to Max × Unit Cost
```

Example:

```excel
=H2*J2
```

This provides a simple operational estimate of the inventory value associated with a replenishment event.

---

## Important Dataset Consideration

The current simulated receiving transactions increased inventory levels substantially.

As a result, the current dataset may show most or all supplies as `OK` in the initial reorder report.

That is not a problem with the formula.

It means the current simulated inventory is above the minimum thresholds.

To demonstrate the reorder workflow more clearly, a later **scenario/stress-test dataset** can intentionally introduce additional simulated usage so that selected supplies fall below their Min levels.

Those scenario transactions should remain clearly identified as simulated test cases rather than being presented as actual hospital inventory activity.

---

# Data Quality Principles

The project emphasizes data accuracy throughout the workflow.

Examples include:

### Valid Supply IDs

Dropdowns prevent users from entering arbitrary supply identifiers.

### Positive Usage Quantities

Usage quantities must be greater than zero.

### Barcode Reference

Barcodes are maintained in the master inventory table and referenced by receiving records.

### QA Status

Receiving and usage transactions include QA status fields.

### Physical Count Validation

Physical counts are kept separate from system quantities.

### Reconciliation

System quantities are compared against physical counts rather than assumed to be accurate.

### Discrepancy Detection

Variance and variance percentage calculations identify records requiring investigation.

### Min/Max Monitoring

Inventory levels are compared against defined replenishment thresholds.

---

# Project Architecture

The overall data flow is:

```text
                    ┌──────────────────┐
                    │ Inventory_Master │
                    │ Reference Data   │
                    └────────┬─────────┘
                             │
            ┌────────────────┼────────────────┐
            ↓                ↓                ↓
    ┌───────────────┐ ┌──────────────┐ ┌───────────────┐
    │ Receiving_Log │ │  Usage_Log   │ │ Physical_Count│
    │   Receipts    │ │   Usage      │ │   Observation │
    └───────┬───────┘ └──────┬───────┘ └───────┬───────┘
            │                │                 │
            └────────────────┼─────────────────┘
                             ↓
                   ┌──────────────────┐
                   │  Reconciliation  │
                   │ System vs Actual │
                   └────────┬─────────┘
                            ↓
                   ┌──────────────────┐
                   │ Reorder_Report   │
                   │ Min / Max / Cost │
                   └────────┬─────────┘
                            ↓
                   ┌──────────────────┐
                   │   QA_Checks      │
                   │   Dashboard      │
                   └──────────────────┘
```

---

# Skills Demonstrated

This project demonstrates practical experience with:

### Inventory & Operations

* Inventory tracking
* Receiving records
* Usage/issuance records
* Physical inventory counts
* Point-of-use inventory concepts
* Min/Max inventory controls
* Replenishment calculations
* Inventory reconciliation
* Stock-level monitoring

### Data Accuracy & Quality

* Data validation
* Record verification
* Barcode validation
* QA workflows
* Discrepancy detection
* Variance analysis
* Exception identification
* Data integrity

### Quantitative Analysis

* Basic inventory calculations
* Variance calculations
* Percentage variance
* Cost calculations
* Threshold-based classification
* Replenishment quantities

### Spreadsheet Skills

* Excel
* VLOOKUP
* SUMIF
* IF / IFERROR
* Data validation
* Conditional formatting
* Operational reporting

---

# Project Limitations

This is a **portfolio simulation**, not a production healthcare inventory management system.

It does not attempt to reproduce:

* A hospital's actual inventory policies
* Actual hospital supply-chain software
* Real purchase orders
* Real supplier relationships
* Actual barcode standards
* Actual par-level policies
* Actual hospital reconciliation thresholds
* Real patient-care data
* Confidential hospital information

Any thresholds, quantities, suppliers, staff IDs, and inventory transactions are fictional and exist solely to demonstrate the workflow.

---

# Current Project Status

Completed through **Step 5**:

* [x] Step 1 — Inventory Master
* [x] Step 2 — Receiving Log
* [x] Step 3 — Usage Log
* [x] Step 4 — Reconciliation
* [x] Step 5 — Physical Counts

Planned future components:

* [ ] Step 6 — Min/Max Reorder Report
* [ ] Step 7 — QA Checks
* [ ] Step 8 — Dashboard
* [ ] Final data-quality review
* [ ] Final formatting and documentation
* [ ] Portfolio screenshots
* [ ] GitHub project documentation
* [ ] Resume project entry
* [ ] Interview explanation

---

# Portfolio Objective

The project is intended to demonstrate that the creator can translate quantitative and data-analysis skills into an operational environment involving:

**accuracy → inventory records → physical verification → reconciliation → exception handling → replenishment → reporting**

The project complements an Applied & Computational Mathematics background by showing practical application of quantitative reasoning, data validation, documentation, and spreadsheet-based operational analysis to a healthcare supply-chain scenario.
