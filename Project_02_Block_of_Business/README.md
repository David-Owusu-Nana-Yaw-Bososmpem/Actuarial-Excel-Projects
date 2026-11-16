# Project 2: Annuity Block of Business Modelling and Portfolio Analysis

## 1. Project Overview

This project extends the Annuity Illustration Tool developed in Project 1 by applying the annuity payment calculations to a portfolio of policyholders, referred to as a Block of Business.

The objective was to develop an Excel-based model that processes multiple annuity policies, calculates their expected future payments and consolidates the results into portfolio-level summaries.

Rather than illustrating the expected payments for one annuitant at a time, the model applies the existing annuity calculation framework to a larger group of clients with different ages, annuity plans and payment arrangements.

The project demonstrates how individual policy-level calculations can be aggregated to support portfolio analysis and actuarial reporting.

## 2. Business Problem

An insurance company manages a portfolio of annuity contracts involving multiple policyholders.

Each policyholder may have a different annuity plan, current age, payment commencement age and payment duration.

Calculating and monitoring the expected payments for every policy individually can become inefficient as the number of policies increases.

The task was to extend the existing annuity illustration model to accommodate a block of business and provide consolidated information about expected payments across the portfolio.

## 3. Project Objectives

The main objectives were to:

- Extend the annuity illustration model developed in Project 1.
- Incorporate multiple policyholders into a single modelling framework.
- Process different annuity plans and their corresponding input requirements.
- Calculate projected annuity payments for individual policyholders.
- Consolidate individual payment schedules into a portfolio-level payment schedule.
- Calculate total payments by annuity plan.
- Analyse projected payments by payment year.
- Present portfolio results through summary tables and graphical visualisations.

## 4. Dataset and Portfolio Structure

The workbook contains a Block of Business worksheet with 100 client records.

Each record includes the following information:

| Variable | Description |
|---|---|
| Client Name | Identifier assigned to each policyholder. |
| Plan Number | Selected annuity product, numbered 1 to 7. |
| Current Age | Current age of the policyholder. |
| Age Payments Begin | Age at which annuity payments commence, where applicable. |
| Years of Payment | Selected payment duration, where applicable. |

The dataset contains policyholders with different annuity plans and payment commencement arrangements.

This variation allows the model to calculate payments under different annuity structures within the same portfolio.

## 5. Workbook Structure

The workbook consists of six worksheets.

### A. Illustration Tool

This worksheet retains the individual annuity illustration functionality developed in Project 1.

It allows the model to retrieve projected payment amounts based on the selected annuity plan and relevant policyholder information.

### B. Plan Calculations

This worksheet contains the calculation logic for the seven annuity plans.

It determines the expected payment for each relevant age and payment period.

The calculations are linked to the annuity inputs and are used as the foundation for generating individual policy payment schedules.

### C. Block of Business

This worksheet contains the 100 policyholder records used in the portfolio model.

It provides the information required to determine the annuity plan, current age, payment commencement age and payment duration for each client.

### D. All Payments

This worksheet consolidates the projected annuity payments for the individual policyholders.

It organises the results by payment year, with individual clients represented in separate columns.

The worksheet provides a consolidated view of expected payments across the portfolio and allows payments to be analysed across different policyholders and annuity products.

### E. Summaries

This worksheet presents consolidated portfolio information in two main sections.

**Total Payment by Plan**

Calculates the total projected payments associated with each of the seven annuity plans.

**Total Payment by Year**

Summarises the expected payments for each annuity plan across the different payment years and calculates the overall payment amount for each year.

This provides a basis for examining the distribution of expected annuity payments over time.

### F. Graphs

This worksheet contains three charts that visually present the portfolio results.

The charts complement the summary tables and provide a graphical view of the projected payment information.

## 6. Methodology and Excel Implementation

The model follows a structured calculation process.

**Step 1: Policyholder data input**

The Block of Business worksheet provides the individual policyholder information, including the selected annuity plan and relevant age and payment details.

**Step 2: Annuity payment calculations**

The Plan Calculations worksheet applies the relevant calculation logic for each of the seven annuity plans.

Each plan has its own payment conditions and calculation formula.

**Step 3: Individual payment projection**

The Illustration Tool uses the selected annuity plan to retrieve the corresponding payment amounts.

The `HLOOKUP` function connects the selected plan to the payment calculation table.

**Step 4: Portfolio payment consolidation**

The All Payments worksheet organises the projected payments for individual policyholders into a consolidated payment schedule.

**Step 5: Portfolio summarisation**

The Summaries worksheet aggregates the projected payments by annuity plan and payment year.

**Step 6: Graphical presentation**

The Graphs worksheet presents the portfolio results through charts to make the projected payment patterns easier to interpret.

### Key Excel functions and features

**HLOOKUP**

Used to retrieve the appropriate annuity payment schedule based on the selected plan number.

**SUMIF**

Used to aggregate payments according to specified criteria.

For example, it is used to calculate total payments associated with individual annuity plans.

**SUM**

Used to calculate aggregate payment amounts across payment periods and policyholders.

**IF and AND**

Used in the plan calculation engine to determine whether the conditions for making a payment have been satisfied.

**ISEVEN and ISODD**

Used in the biennial annuity calculation to identify the appropriate payment years for Plan 7.

**Named Ranges**

The named range `plan_payment` connects the annuity plan calculation table to the illustration calculator.

**Excel Charts**

Used to present consolidated payment information graphically.

## 7. Portfolio Analysis

The project introduces two important dimensions of portfolio analysis.

### A. Payments by Annuity Plan

The model calculates the total projected payments associated with each annuity product.

This allows the portfolio to be examined according to its different payment structures.

For example, the company can compare the aggregate payment obligations associated with fixed-payment annuities, increasing annuities and biennial annuities.

### B. Payments by Year

The model also consolidates expected payments across different payment years.

This provides information about how the portfolio's projected payment obligations are distributed over time.

The analysis can help identify years with relatively high or low expected annuity payments.

## 8. Quality Assurance and Reasonability Checks

The reliability of the portfolio model depends on the accuracy of the underlying policy-level calculations and the aggregation process.

Recommended verification procedures include:

- Confirming that each client is assigned the correct annuity plan.
- Checking that the current age, payment commencement age and payment duration are correctly captured.
- Comparing selected individual policy projections against the original Annuity Illustration Tool.
- Confirming that individual client payments are included in the consolidated payment schedule.
- Checking that total payments by annuity plan reconcile with the underlying client payment records.
- Confirming that annual portfolio totals reconcile with the sum of payments across all clients.
- Checking that the summary tables and charts reflect the underlying payment data.
- Testing policyholders with different annuity plans and payment commencement arrangements.

These checks help identify errors in individual calculations, data processing and portfolio aggregation.

## 9. Technical Skills Demonstrated

- Microsoft Excel
- Advanced Excel formulas
- Annuity payment modelling
- Block of business analysis
- Portfolio-level calculations
- HLOOKUP and named ranges
- SUMIF and SUM functions
- Conditional logic using IF and AND
- Data aggregation and reconciliation
- Financial projection
- Excel charts and visualisation
- Spreadsheet modelling and organisation
- Actuarial data analysis

## 10. Actuarial Relevance

This project demonstrates the application of individual insurance benefit calculations to a portfolio of annuity contracts.

The ability to process policy-level information and consolidate projected payments is relevant to several actuarial functions, including:

- Annuity portfolio analysis.
- Insurance benefit projections.
- Cash flow analysis.
- Actuarial reporting.
- Policy-level data processing.
- Portfolio reconciliation.
- Financial planning and forecasting.

It also provides practical experience in organising large sets of policyholder information and translating individual policy calculations into consolidated portfolio results.

## 11. Limitations and Further Improvements

Potential improvements include:

- Automating the processing of additional policyholders without manually extending formulas.
- Introducing stronger input validation for incomplete or inconsistent policy records.
- Adding automated reconciliation checks between individual policy payments and portfolio summaries.
- Developing interactive filters for analysing payments by annuity plan.
- Incorporating additional portfolio metrics, including the number of policies under each plan.
- Adding sensitivity analysis to examine how changes in policyholder characteristics and payment assumptions affect projected obligations.
- Improving the graphical presentation of portfolio payment trends.

## 12. Project Files

| File | Description |
|---|---|
| `Annuity_Block_of_Business_Model.xlsx` | Excel workbook containing the individual annuity calculator, block of business, consolidated payment schedules, summary tables and charts. |

## 13. Conclusion

This project extended an individual annuity illustration model into a portfolio-level modelling framework.

By applying the existing annuity calculation logic to 100 policyholders, I developed an Excel workbook that organises individual payment projections, consolidates expected payments and summarises portfolio obligations by annuity plan and payment year.

The project strengthened my practical understanding of annuity modelling, portfolio analysis, data aggregation and actuarial reporting using Microsoft Excel.
