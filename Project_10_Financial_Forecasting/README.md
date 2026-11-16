# Automated Insurance Income Statement Forecasting and Scenario Analysis | Excel VBA

## Project Overview

This project extends the insurance income statement projection model developed in Project 9.

The objective was to automate the calculation of cumulative net income for 100 different financial forecasting scenarios under four interest-rate assumptions: 2%, 4%, 6% and 8%.

Using the existing income statement projection tool and VBA programming, I developed a solution designed to process different combinations of annual change assumptions and calculate the resulting net income for each scenario.

In total, the model is designed to perform 400 net income calculations, reducing the need for repetitive manual forecasting.

## Project Objectives

- Automate income statement projections across 100 different scenarios.
- Evaluate projected net income under four interest-rate assumptions.
- Apply nine different annual change assumptions to each scenario.
- Incorporate interest-rate adjustments into selected financial assumptions.
- Calculate cumulative net income over the five-year period (2021–2025).
- Use VBA loops to automate repetitive calculations and populate the results.

## Dataset Description

The supplied forecasting dataset contains 100 scenarios, each with nine separate annual change assumptions.

| Forecast Input | Description |
|---|---|
| Gross Premiums | Annual change assumption for gross premiums |
| Premiums Ceded to Reinsurers | Annual change assumption for ceded premiums |
| Investment Income | Annual change assumption for investment income |
| Gross Claims | Annual change assumption for gross claims |
| Benefits | Annual change assumption for benefits |
| Equipment Expenses | Annual change assumption for equipment expenses |
| Claims and Expenses Ceded to Reinsurers | Annual change assumption for ceded claims and expenses |
| Acquisition Costs Amortization | Annual change assumption for acquisition costs |
| Change in Insurance Liabilities | Annual change assumption for insurance liabilities |

## Methodology and Implementation

### 1. Income Statement Projection

The model uses the 2021 income statement as its starting point and projects financial performance through 2025.

For each scenario, the nine forecast assumptions are transferred to the assumption tool and applied to their corresponding income statement items.

The model calculates:

- Projected revenue.
- Projected expenses.
- Annual net income.
- Cumulative net income from 2021 to 2025.

### 2. Interest Rate Scenario Analysis

Each forecasting scenario is evaluated under four different interest-rate assumptions:

- 2%
- 4%
- 6%
- 8%

The interest rate determines whether specific annual change assumptions are adjusted.

The model incorporates the following scenario adjustments:

| Interest Rate | Financial Item | Adjustment |
|---|---|---:|
| 2%–4% | Investment Income | Decrease by 0.25% |
| 5%–7% | Investment Income | Increase by 0.625% |
| 2%–4% | Gross Claims | Increase by 1% |
| 5%–7% | Gross Claims | Decrease by 0.7% |
| 2%–4% | Change in Insurance Liabilities | Increase by 1.5% |
| 2%–4% | Equipment Expenses | Increase by 2% |

These adjustments are multiplicative and are applied to the original annual change assumptions.

The model therefore evaluates how different interest-rate environments influence the projected financial performance of the insurance company.

### 3. VBA Automation

VBA programming is used to automate the processing of the 100 forecasting scenarios.

The automation is designed to:

1. Read the forecast assumptions for each scenario.
2. Transfer the assumptions into the existing income statement projection tool.
3. Apply the selected interest-rate assumption.
4. Calculate the projected income statement for 2021–2025.
5. Obtain the cumulative net income.
6. Store the calculated result in the corresponding output column.
7. Repeat the process for all 100 scenarios and four interest rates.

The use of VBA loops makes it possible to process multiple scenarios efficiently without manually repeating the calculations.

## Expected Output

The `Forecasting_Data` worksheet contains four output columns corresponding to the interest-rate assumptions.

| Output Column | Interest Rate |
|---|---:|
| INT=0.02 | 2% |
| INT=0.04 | 4% |
| INT=0.06 | 6% |
| INT=0.08 | 8% |

Each output represents the cumulative projected net income for the corresponding scenario and interest rate.

The completed model is designed to produce 400 calculations across the 100 scenarios.

## Workbook Structure

| Worksheet | Description |
|---|---|
| Assumptions | Contains the annual change assumptions and interest-rate input. |
| Income Statement Projection | Calculates projected revenue, expenses and net income. |
| Forecasting_Data | Contains the 100 scenarios and four corresponding output columns. |
| Scenarios | Stores the interest-rate adjustments applied to selected assumptions. |

## Tools and Techniques Used

- Microsoft Excel
- Visual Basic for Applications (VBA)
- VBA loops
- Excel formulas
- Conditional statements
- Dynamic cell referencing
- Financial forecasting
- Scenario analysis
- Automated data processing

## Key Skills Demonstrated

- Insurance financial modelling
- Income statement projections
- Net income forecasting
- Interest-rate scenario analysis
- Assumption-based modelling
- VBA automation
- Repetitive calculation automation
- Financial risk assessment

## Actuarial Relevance

Scenario analysis is an important component of actuarial work, particularly when evaluating how changes in economic and financial assumptions affect an insurance company's financial position.

This project demonstrates the application of financial modelling and automation to evaluate multiple possible outcomes under different interest-rate environments.

The skills developed are relevant to actuarial analysis, insurance financial planning, risk management, financial forecasting and performance assessment.

## Repository Contents

- `Financial_Forecasting.xlsm` – Macro-enabled Excel workbook containing the income statement projection model, forecasting dataset, scenario assumptions and VBA automation.

