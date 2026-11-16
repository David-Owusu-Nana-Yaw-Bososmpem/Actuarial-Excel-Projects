# IBNR Claims Reserving Calculator Using the Completion Factor Method

## Project Overview

This project involves developing an Excel-based actuarial tool to estimate Ultimate Expected Claims and Incurred But Not Reported (IBNR) Reserves for a group health insurance portfolio using the Completion Factor Method.

The calculator estimates claims that have been incurred but not yet reported by applying predetermined completion factors to reported claims. It automatically selects the appropriate factors based on employee count and claim type, eliminating the need for manual formula adjustments.

## Business Problem

Health insurance claims are not always reported immediately after the insured event occurs. Consequently, at a particular valuation date, some incurred claims may remain unreported.

Assuming a six-month claims reporting period and a valuation date of June 30, 2020, the objective was to develop a calculator that estimates:

- Ultimate Expected Claims for January–June 2020.
- The IBNR Reserve required to account for outstanding claims.

## Project Objectives

- Develop an automated Excel-based IBNR calculator.
- Apply completion factors to estimate ultimate expected claims.
- Classify insurance groups based on employee count.
- Incorporate different completion factors for different claim types.
- Calculate IBNR reserves using reported and ultimate expected claims.
- Build a reusable tool that accepts different inputs without manually changing formulas.

## Methodology

The calculator uses completion factors supplied as reference data. These factors vary according to employee category, claim type and development period.

**Employee Categories**

| Category | Employee Count |
|---|---:|
| Small | 1–100 |
| Medium | 101–5,000 |
| Large | 5,001+ |

**Claim Types**
- Dental
- Prescription
- Healthcare Appointment

The calculator performs the following operations:

1. Accepts employee count, claim type and monthly reported claims.
2. Automatically classifies the insurance group.
3. Combines employee category and claim type to generate a lookup key.
4. Retrieves the appropriate completion factors using Excel's VLOOKUP function.
5. Applies the corresponding development factors to monthly reported claims to estimate ultimate expected claims.
6. Calculates the IBNR Reserve by deducting total reported claims from ultimate expected claims.

## Excel Implementation

The project demonstrates the following Excel skills:

| Excel Feature | Application |
|---|---|
| IF and AND | Employee-count classification |
| VLOOKUP | Retrieval of completion factors |
| SUM | Aggregation of reported claims |
| Data Validation | Claim-type dropdown selection |
| Cell Referencing | Connecting inputs and calculations |
| Concatenation | Generating lookup keys |
| Absolute References | Maintaining consistent formula references |

## Workbook Structure

The Excel solution contains three worksheets:

| Worksheet | Description |
|---|---|
| Calculator | Main interface for entering inputs and calculating ultimate claims and IBNR reserves. |
| Completion_Factors | Reference table containing completion factors for different employee categories and claim types. |
| Reported_Claims | Example claims data used to test the calculator. |

## Illustrative Results

**Valuation Date:** June 30, 2020

**Employee Count:** 10,000

**Employee Category:** Large

**Claim Type:** Prescription

| Metric | Result |
|---|---:|
| Total Reported Claims | $387,056.00 |
| Ultimate Expected Claims | $434,436.39 |
| **IBNR Reserve** | **$47,380.39** |

**Interpretation:** The model estimates ultimate claims of $434,436.39, compared with $387,056.00 already reported. The resulting IBNR reserve is $47,380.39, representing estimated additional claims expected beyond those already reported.

## Actuarial Applications

This project demonstrates practical applications of claims reserving and insurance modelling, including:

- Estimation of ultimate claims from incomplete reported claims data.
- Calculation of outstanding claims reserves.
- Analysis of different claims development patterns.
- Automation of actuarial reserving calculations.
- Support for insurance financial reporting and claims management.

## Assumptions and Limitations

- The valuation date is June 30, 2020.
- Claims are assumed to be reported within six months.
- Completion factors are supplied as predetermined reference data.
- Completion factors vary by employee category and claim type.
- The model does not independently estimate completion factors from historical claims-development data.

Further enhancements could include historical claims-development analysis, additional input validation, sensitivity analysis and automated reporting.

## Technologies and Tools

- Microsoft Excel
- Excel formulas and functions
- Data Validation
- Lookup and reference functions
- Completion Factor Method
- Actuarial Claims Reserving

## Conclusion

This project provided practical experience in applying the Completion Factor Method to estimate ultimate expected claims and IBNR reserves. It strengthened my understanding of claims development, actuarial reserving and the application of Excel in developing automated insurance modelling tools.

---

**Project Area:** Actuarial Science | Health Insurance | Claims Reserving | IBNR Modelling

**Tool:** Microsoft Excel
