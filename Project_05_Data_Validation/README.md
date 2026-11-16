# Actuarial Data Validation and Continuity Checks

## Project Overview

This project involves developing an Excel-based actuarial data validation tool for EA Insurance Company. The tool performs continuity checks by reconciling primary financial data points obtained from actuarial software against their underlying granular components.

The objective is to verify the accuracy, consistency and integrity of financial data during the company's quarterly reporting process.

The model covers expense, revenue and net income reconciliations for the prior period (M0: October–December 2022) and current period (Q1: January–March 2023).

## Project Objectives

- Reconcile primary financial data points against their granular components.
- Validate expense and revenue figures across reporting periods.
- Develop continuity checks for net income.
- Identify discrepancies between reported totals and their underlying components.
- Automate data aggregation using Excel formulas.
- Present reconciliation results in a clear format for actuarial review.

## Methodology

The project uses a continuity-check framework in which primary data points are compared with the corresponding granular data points.

The model is structured around nine continuity groups:

| Continuity Group | Reconciliation Purpose |
|---|---|
| CONT_GROUP_00 | Prior-period total expenses |
| CONT_GROUP_01 | Current-period expense components |
| CONT_GROUP_02 | Current-period total expenses |
| CONT_GROUP_03 | Change in expenses |
| CONT_GROUP_04 | Prior-period total revenue |
| CONT_GROUP_05 | Current-period revenue components |
| CONT_GROUP_06 | Current-period total revenue |
| CONT_GROUP_07 | Prior-period net income |
| CONT_GROUP_08 | Current-period net income |

Each continuity group compares the aggregated primary value with its corresponding granular value. A successful reconciliation returns TRUE, while a mismatch returns FALSE.

## Excel Implementation

### 1. Data Extraction and Aggregation
The `Primary Values` and `Granular Values` worksheets retrieve and aggregate the relevant financial figures from the `Actuarial Data` worksheet using formulas linked to the primary and granular data point definitions.

### 2. Continuity Group Mapping
The `Continuity Group for Primary` and `Continuity Group for Granular` worksheets organize the calculated values into their respective continuity groups.

### 3. Reconciliation Checks
The `Continuity Group` worksheet compares the aggregated primary and granular values using SUMIFS formulas. It also calculates the differences and displays the reconciliation status.

### 4. Revenue and Net Income Validation
Additional data point definitions and formulas were developed to incorporate total revenue and net income into the reconciliation framework.

Revenue was reconciled against premiums, commission, investment earnings and other revenue. Net income was also incorporated into the validation process.

### 5. Summary and Conditional Formatting
VLOOKUP formulas generate summary statements for the continuity groups. Conditional formatting highlights successful checks in green and unsuccessful checks in red, making discrepancies easier to identify.

Numerical values are also formatted for readability.

## Key Results

The completed workbook successfully reconciles all nine continuity groups.

| Financial Reconciliation | Primary Value | Granular Value | Result |
|---|---:|---:|---|
| Prior-period expenses | 58,500,000 | 58,500,000 | TRUE |
| Current-period expenses | 50,250,000 | 50,250,000 | TRUE |
| Change in expenses | 6,500,000 | 6,500,000 | TRUE |
| Prior-period revenue | 62,050,000 | 62,050,000 | TRUE |
| Current-period revenue | 57,120,000 | 57,120,000 | TRUE |
| Prior-period net income | 3,550,000 | 3,550,000 | TRUE |
| Current-period net income | 6,870,000 | 6,870,000 | TRUE |

**Validation outcome:** All nine continuity groups return TRUE, indicating that the primary and granular values reconcile in the completed workbook.

## Workbook Structure

| Worksheet | Purpose |
|---|---|
| Continuity Group | Main reconciliation dashboard and summary points |
| Continuity Group for Primary | Aggregated primary values by continuity group |
| Continuity Group for Granular | Aggregated granular values by continuity group |
| Primary Values | Calculated primary data point values |
| Granular Values | Calculated granular data point values |
| Actuarial Data | Source financial data from actuarial software |
| Primary Data Point Formulas | Definitions of primary data point calculations |
| Granular Data Point Formulas | Definitions of granular data point calculations |

## Tools and Techniques

- Microsoft Excel
- SUMIFS
- SUMIF
- IF
- VLOOKUP
- Excel conditional formatting
- Financial data reconciliation
- Actuarial data validation
- Formula-based aggregation and reporting

## Actuarial Relevance

Data validation and reconciliation are essential components of actuarial work, particularly in financial reporting, reserving, valuation and risk management.

This project demonstrates practical experience in developing structured validation checks, reconciling financial data, identifying inconsistencies and presenting results in a format that supports actuarial review and financial reporting.

The skills demonstrated are relevant to actuarial analyst, actuarial reporting, risk and financial analysis roles.
