# Insurance Income Statement Projection and Scenario Analysis | Microsoft Excel

## Project Overview

This project involved developing an Excel-based financial modelling tool to project an insurance company's income statement and forecast net income over a five-year period (2021–2025).

Using the company's completed 2021 income statement as the starting point, I developed a dynamic projection model that allows users to modify annual change assumptions for different revenue and expense items.

The model also incorporates interest-rate scenarios that automatically adjust selected assumptions, allowing users to assess how changes in economic conditions can influence the company's projected financial performance.

## Project Objectives

- Develop a five-year insurance income statement projection model.
- Forecast revenue, expenses and net income from 2022 to 2025.
- Allow users to modify annual percentage change assumptions.
- Incorporate increase, decrease and no-change options for individual income statement items.
- Introduce interest-rate scenarios that automatically adjust selected assumptions.
- Evaluate the impact of different assumptions on projected net income.

## Model Assumptions

The model incorporates annual change assumptions for the following income statement items:

### Revenue
- Gross Premiums
- Premiums Ceded to Reinsurers
- Investment Income

### Expenses
- Gross Claims
- Benefits
- Equipment Expenses
- Claims and Expenses Ceded to Reinsurers
- Acquisition Costs Amortization
- Change in Insurance Liabilities

Users can specify the direction of change (Increase, Decrease or No Change) and the corresponding percentage for each item.

The model uses the assumptions provided in the assignment as its default reference, while allowing users to modify them for alternative projection scenarios.

## Methodology and Implementation

### 1. Income Statement Projection

The completed 2021 income statement serves as the base year for projecting financial performance from 2022 to 2025.

Each revenue and expense item is projected independently using its corresponding annual change assumption.

The model automatically carries forward each year's projected value to calculate the following year's amount.

Total Revenue and Total Expenses are calculated by aggregating their respective components.

Net Income is then determined from the difference between Total Revenue and Total Expenses.

### 2. Dynamic Assumption Tool

The Assumptions worksheet allows users to modify the direction and percentage of annual change for each income statement item.

The model supports three possible directions:

- **Increase:** Applies the specified percentage increase.
- **Decrease:** Applies the specified percentage decrease.
- **No Change:** Retains the previous year's value.

This makes the projection flexible and eliminates the need to manually modify individual projection formulas when testing different assumptions.

### 3. Interest Rate Scenario Analysis

The model incorporates an interest-rate input that automatically adjusts selected annual change assumptions depending on the interest-rate range.

The scenarios implemented are:

| Interest Rate | Income Statement Item | Adjustment |
|---|---|---:|
| 2%–4% | Investment Income | Decrease by 0.25% |
| 5%–7% | Investment Income | Increase by 0.625% |
| 2%–4% | Gross Claims | Increase by 1% |
| 5%–7% | Gross Claims | Decrease by 0.7% |
| 2%–4% | Change in Insurance Liabilities | Increase by 1.5% |
| 2%–4% | Equipment Expenses | Increase by 2% |

These adjustments are multiplicative and are applied to the original annual change assumptions.

For example, when the interest rate falls within the 5%–7% range, the Investment Income annual change assumption is adjusted upwards by multiplying the original assumption by 1.00625.

The model automatically applies the relevant scenario adjustments without requiring users to manually change the underlying assumptions.

## Workbook Structure

The workbook contains three worksheets:

| Worksheet | Description |
|---|---|
| Assumptions | Contains editable directions and annual percentage change assumptions, including the interest-rate input. |
| Income Statement Projection | Contains the historical 2021 income statement and projected financial results for 2022–2025. |
| Scenarios | Contains interest-rate ranges and their corresponding adjustments to selected assumptions. |

## Excel Functions and Techniques Used

- `IF` – Applies increase, decrease and no-change conditions.
- `AND` – Checks whether an interest rate falls within a specified range.
- `SUM` – Aggregates revenue and expense components.
- Cell referencing – Connects assumptions, scenarios and income statement projections.
- Nested conditional formulas – Automates the application of multiple scenario conditions.

## Key Skills Demonstrated

- Financial modelling and forecasting
- Insurance income statement analysis
- Net income projection
- Assumption-based modelling
- Scenario and sensitivity analysis
- Excel formula automation
- Revenue and expense forecasting
- Financial performance assessment

## Actuarial Relevance

Financial projections and scenario analysis are important in actuarial work, particularly when assessing an insurance company's financial performance, profitability and exposure to changing economic conditions.

This project demonstrates the ability to develop flexible financial models, incorporate assumptions into projections and evaluate how changes in interest rates can influence key financial indicators.

The skills developed are relevant to actuarial analysis, financial reporting, insurance risk management and financial planning.

## Repository Contents

- `Income_Statement_Projectio.xlsx` – Excel workbook containing the assumption tool, income statement projections and interest-rate scenario analysis.

