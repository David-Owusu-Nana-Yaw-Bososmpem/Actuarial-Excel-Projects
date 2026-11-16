# Actuarial Data Validation and Reporting Automation | Excel VBA

## Project Overview

This project extends an actuarial data validation tool by introducing Excel VBA automation for quarterly data updates and financial reasonability checks.

Developed for EA Insurance Company, the tool is designed to reduce manual intervention in the quarterly actuarial reporting process by automating the retrieval of financial data, updating actuarial data tables and assessing whether key financial values fall within predefined expected ranges.

The solution builds on the continuity-check framework developed in Project 5.

## Project Objectives

- Automate the updating of actuarial data for new reporting periods.
- Retrieve prior-period and current-period financial data from an external Excel workbook.
- Use VBA to clear and refresh existing data.
- Update data according to the selected valuation period.
- Perform reasonability checks on total expenses, total revenue and net income.
- Highlight values that fall outside expected ranges.
- Improve the usability and efficiency of the actuarial validation process.

## Methodology and Implementation

### 1. Quarterly Data Update Automation

The `Actuarial Data` worksheet contains financial information for different reporting periods.

VBA automation is incorporated to:
- Clear existing actuarial data before a refresh.
- Open an external workbook containing the new reporting-period data.
- Retrieve the relevant prior-period and current-period values.
- Populate the actuarial data table.
- Close the external workbook after retrieving the required information.

The `ValuationPeriodData` component supports the selection of the appropriate reporting-period columns, allowing the tool to accommodate different quarters.

### 2. Dynamic Data Aggregation

The `Primary Values` and `Granular Values` worksheets use formulas that reference the selected valuation period to retrieve the appropriate financial figures.

This allows the existing continuity-check framework to work with updated reporting-period data without manually rebuilding the entire model.

### 3. Financial Reasonability Checks

The `Checks & Inferences` worksheet evaluates current-period values for three financial indicators against predefined thresholds.

| Financial Indicator | Expected Range |
|---|---|
| Total Expenses | 45,000,000–55,000,000 |
| Total Revenue | 50,000,000–60,000,000 |
| Net Income | 1,000,000–5,500,000 |

The checks classify each value as either within or outside its expected range.

### 4. Automated Validation and Reporting

The VBA macros support the updating and clearing of reasonability-check results.

The tool applies the following validation outcomes:

- **Within range:** Displays "Value is within the expected range." and highlights the amount in green.
- **Outside range:** Displays "Value is not within the expected range. Please contact the department lead." and highlights the amount in red.

This provides a straightforward way for analysts and reviewers to identify financial values that require further investigation.

## Key VBA Components

| VBA Component | Purpose |
|---|---|
| `ValuationPeriodData` | Supports valuation-period selection and data-update functionality |
| `ClearACTdata` | Clears existing actuarial data before updating |
| `re_check` | Performs the reasonability-check process |
| `ClearREchecks` | Clears existing reasonability-check results |

## Workbook Structure

| Worksheet | Purpose |
|---|---|
| Continuity Group | Main financial reconciliation checks |
| Continuity Group for Primary | Aggregated primary data points |
| Continuity Group for Granular | Aggregated granular data points |
| Primary Values | Primary financial data calculations |
| Granular Values | Granular financial data calculations |
| Actuarial Data | Financial data used by the validation model |
| Checks & Inferences | Financial reasonability checks and explanatory notes |
| Provided Data | Supplied reporting-period data |
| Primary Data Point Formulas | Definitions of primary data point calculations |
| Granular Data Point Formulas | Definitions of granular data point calculations |

## Tools and Techniques

- Microsoft Excel
- Visual Basic for Applications (VBA)
- Excel formulas
- Loops and automated data processing
- Dynamic reporting-period selection
- Financial data validation
- Reasonability checks
- Conditional formatting
- Actuarial data reconciliation

## Actuarial Relevance

Actuarial data validation is essential to the reliability of financial reporting, reserving, valuation and risk management processes.

This project demonstrates practical experience in automating recurring actuarial tasks, maintaining data consistency across reporting periods and applying reasonability checks to identify potentially unusual financial values.

It highlights skills relevant to actuarial reporting, financial analysis, risk management and actuarial operations, particularly where accuracy, efficiency and reliable data processes are important.
