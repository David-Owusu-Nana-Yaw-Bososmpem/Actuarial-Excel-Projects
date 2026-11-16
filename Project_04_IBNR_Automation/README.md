# IBNR Claims Reserving and Portfolio Analysis

## Project Overview

This project extends an individual IBNR (Incurred But Not Reported) claims reserving calculator into a multi-company reserving and portfolio analysis tool using Microsoft Excel and VBA.

The model estimates outstanding incurred claims for 30 insurance companies using the Completion Factor Method. It also uses PivotTable analysis to summarize claims experience across different coverage types.

## Project Objectives

- Automate IBNR reserve calculations for multiple insurance companies.
- Apply completion factors based on company size and coverage type.
- Analyse reported claims across different coverage categories.
- Use Excel VBA to improve the efficiency of repetitive actuarial calculations.
- Apply PivotTables to summarize portfolio-level claims information.

## Methodology

The project builds on the Completion Factor Method developed in Project 3.

For each company, the model uses:
- Company size to determine its employee category.
- Coverage type to identify the appropriate completion factors.
- Monthly reported claims from January to June 2020.
- Completion factors corresponding to the reporting period of each month's claims.

The model estimates ultimate claims and calculates the IBNR reserve by comparing estimated ultimate claims with claims already reported.

## Excel Implementation

### 1. IBNR Calculator
The `Calculator` worksheet contains the underlying reserving model. It determines the employee category, retrieves the appropriate completion factors and calculates ultimate claims and the IBNR reserve.

### 2. VBA Automation
The workbook includes VBA automation for applying the reserving process across multiple companies, reducing the need to calculate each company's reserve manually.

### 3. Company Data
The `Company_Data` worksheet contains information for 30 companies, including:
- Company identifier
- Coverage type
- Company size
- Monthly reported claims
- Calculated IBNR reserve

### 4. Completion Factors
The `Completion_Factors` worksheet contains the completion factors used in the reserving calculations for different company-size and coverage-type combinations.

### 5. PivotTable Analysis
The `Pivot_Table` worksheet contains portfolio summaries, including:
- Distribution of companies across coverage types.
- Average monthly reported claims by coverage type.

## Workbook Structure

| Worksheet | Description |
|---|---|
| Calculator | Individual IBNR calculation model |
| Company_Data | Dataset containing 30 companies and their reserve results |
| Completion_Factors | Completion factors by company size and coverage type |
| Reported_Claims | Historical reported claims data |
| Pivot_Table | Portfolio-level PivotTable analysis |

## Tools and Techniques

- Microsoft Excel
- Excel formulas and lookup functions
- Visual Basic for Applications (VBA)
- PivotTables
- Claims reserving
- Completion Factor Method
- Portfolio analysis

## Actuarial Relevance

This project demonstrates practical skills relevant to actuarial work, particularly in general insurance and health insurance claims reserving.

It highlights the ability to automate repetitive calculations, estimate outstanding claims liabilities, organize claims data and summarize portfolio experience to support actuarial analysis and decision-making.

## Limitations

The model uses supplied completion factors and historical claims information. Its estimates therefore depend on the reliability and suitability of those assumptions.

The current workbook also contains an incomplete PivotTable section for maximum IBNR reserve by coverage type. This analysis should be completed before presenting the workbook as satisfying every requirement of the original assignment.
