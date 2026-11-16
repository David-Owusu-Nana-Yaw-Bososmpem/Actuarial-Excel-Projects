# Project 1: Annuity Illustration Tool in Microsoft Excel

## 1. Project Overview

This project involved developing an Excel-based Annuity Illustration Tool for a hypothetical insurance company that offers different annuity products to individuals aged 40 and above.

The objective was to create an intuitive and professional tool that allows insurance agents to generate annual annuity payment illustrations based on the selected annuity plan and the annuitant's information.

The calculator determines the expected payment at each relevant age and presents the projected payment schedule, total payments and age of the last payment.

The project was designed to simulate a practical task that an actuarial analyst could encounter in an insurance company, particularly in insurance product modelling and benefit projection.

## 2. Business Problem

The insurance company offers several annuity products with different payment structures. These include fixed annual payments, increasing payments, age-dependent payments, split payment amounts and payments made at two-year intervals.

Manually calculating these payment schedules for individual clients could be time-consuming and prone to errors.

The task was therefore to develop an Excel-based solution that automates the calculations and enables insurance agents to produce consistent annuity payment illustrations.

## 3. Project Objectives

The main objectives were to:

- Develop an Excel-based annuity illustration calculator.
- Incorporate seven different annuity payment plans.
- Allow agents to select an annuity plan and provide the relevant input information.
- Automatically calculate annual payments based on the selected plan.
- Generate a payment schedule showing the annuitant's age and expected payment.
- Calculate the total projected payments.
- Identify the age of the last projected payment.
- Incorporate logical conditions to distinguish between the different payment structures.

## 4. Annuity Products Modelled

The illustration tool incorporates seven annuity plans.

| Plan | Description                  | Payment structure                                                                         |
| ---- | ---------------------------- | ----------------------------------------------------------------------------------------- |
| 1    | Fixed annual annuity         | $15,000 annually for the selected payment period.                                         |
| 2    | Increasing annuity           | Starts at $25,000, increasing by $1,000 annually, with payments specified through age 94. |
| 3    | Fixed lifetime-style annuity | $7,200 annually from the selected starting age through age 110.                           |
| 4    | Age-dependent annuity        | Annual payment equals the annuitant's age multiplied by $1,000, between ages 60 and 100.  |
| 5    | Two-tier annuity             | $19,000 for the first half of the payment period and $25,000 for the second half.         |
| 6    | Fixed-total annuity          | $100,000 distributed equally across the selected payment period.                          |
| 7    | Biennial annuity             | $30,000 paid every second year, beginning at the selected starting age.                   |

The first five plans formed the initial illustration tool. Plans 6 and 7 were incorporated as additional products in the second phase of the assignment.

## 5. Workbook Structure

The workbook consists of two worksheets.

### A. Illustration Tool

This is the main interface intended for the insurance agent.

It contains:

- Annuity plan selection.
- Annuitant information and plan inputs.
- Projected payment schedule.
- Age of the annuitant for each payment period.
- Expected annual payment.
- Total projected payments.
- Age of the last payment.

The payment schedule is generated using formulas that retrieve the appropriate payment amounts from the Plan Calculations worksheet.

### B. Plan Calculations

This worksheet serves as the calculation engine.

It contains separate calculation columns for all seven annuity plans, with each formula determining the expected payment at a particular age.

The worksheet uses the annuitant's age, selected starting age and payment duration to determine whether a payment should be made and the corresponding amount.

A named range, `plan_payment`, references the plan calculation table. This allows the illustration calculator to retrieve the relevant payment schedule dynamically.

## 6. Methodology and Excel Implementation

The project uses conditional logic, lookup functions and aggregation formulas to automate the annuity calculations.

### Key Excel functions

**IF**

Used to determine whether an annuity payment should be made based on the annuity plan's conditions.

For example, a payment is only generated when the annuitant reaches the specified starting age and remains within the applicable payment period.

**AND**

Used when multiple conditions must be satisfied simultaneously.

This is particularly useful for checking whether the annuitant's age falls within the specified payment period and does not exceed the applicable maximum age.

**OR**

Identified in the project instructions as a useful logical function for evaluating alternative conditions.

**HLOOKUP**

Used in the Illustration Tool to retrieve the expected payment corresponding to the selected annuity plan from the calculation table.

**SUM**

Used to calculate the total projected annuity payments.

**ISEVEN and ISODD**

Used in Plan 7 to determine the appropriate years for biennial payments. These functions help distinguish payment years from non-payment years based on the selected starting age.

### Plan-specific calculation logic

- Plan 1: Applies a fixed annual payment of $15,000 throughout the selected payment period.
- Plan 2: Calculates increasing payments using an initial amount of $25,000 and annual increments of $1,000.
- Plan 3: Applies a fixed annual payment of $7,200 up to the specified maximum age.
- Plan 4: Calculates annual payments as the annuitant's age multiplied by $1,000.
- Plan 5: Uses conditional logic to allocate $19,000 to the first half of the payment period and $25,000 to the second half.
- Plan 6: Divides the total payment obligation of $100,000 equally across the selected payment duration.
- Plan 7: Uses ISEVEN and ISODD to identify the appropriate biennial payment years and allocate $30,000 to each payment year.

## 7. Additional Requirements Incorporated

The second phase of the project introduced two additional annuity products.

**Plan 6: Equal Distribution of a Fixed Total**

The model calculates the annual payment by dividing $100,000 by the selected payment duration.

The intended relationship is:

Annual Payment = $100,000 / Number of Payment Years.

**Plan 7: Biennial Payments**

The model uses the parity of the annuitant's age and selected starting age to determine the payment years.

A payment of $30,000 is assigned to qualifying payment years, while intervening years receive no payment.

## 8. Quality Assurance and Reasonability Checks

The project instructions emphasised the importance of checking calculations before delivering the workbook.

The recommended verification procedures include:

- Confirming that payments do not begin before the selected starting age.
- Checking that payment periods remain within the applicable maximum age.
- Verifying that Plan 4 payments correspond to the annuitant's age multiplied by $1,000.
- Confirming that Plan 5 applies the correct payment amount to each half of the payment period.
- Checking that Plan 6 payments sum to $100,000.
- Confirming that Plan 7 produces alternating payment and non-payment years.
- Testing invalid plan selections and inconsistent age inputs.
- Comparing selected calculator outputs against independently calculated expected results.

These checks are important for maintaining the accuracy and reliability of actuarial calculations.

## 9. Technical Skills Demonstrated

- Microsoft Excel
- Advanced Excel formulas
- Conditional logic and nested IF statements
- AND and OR functions
- HLOOKUP and named ranges
- SUM and payment aggregation
- ISEVEN and ISODD functions
- Annuity payment modelling
- Financial calculations
- Spreadsheet design and organisation
- Actuarial modelling and problem-solving

## 10. Actuarial Relevance

This project demonstrates practical Excel modelling skills applicable to insurance and actuarial operations.

The ability to calculate and project policy benefits is relevant to insurance product administration, benefit illustration, policy servicing and actuarial modelling.

It also demonstrates an understanding of how different payment structures can be translated into mathematical conditions and automated spreadsheet calculations.

The project provides practical experience in building calculation tools that support consistent financial projections and insurance-related decision-making.

## 11. Limitations and Further Improvements

Potential improvements include:

- Adding a dedicated Total Years of Payments output.
- Introducing stronger input validation for invalid or inconsistent ages and plan selections.
- Implementing automated reasonability checks for all seven plans.
- Developing a structured test case register covering normal, boundary and unexpected scenarios.
- Reviewing and independently validating all plan-specific formulas.
- Improving the presentation of the calculator with additional visual summaries.

## 12. Project Files

| File                                  | Description                                                                                |
| ------------------------------------- | ------------------------------------------------------------------------------------------ |
| `Annuity_Illustration_Model.xlsx` | Excel workbook containing the annuity illustration calculator and plan calculation engine. |

## 13. Conclusion

This project provided practical experience in developing an Excel-based annuity illustration tool that incorporates multiple payment structures.

By combining conditional formulas, lookup functions and mathematical calculations, I developed a spreadsheet model designed to automate annuity payment projections and support insurance agents in illustrating expected future benefits.

The project strengthened my practical understanding of Excel-based actuarial modelling, insurance benefit calculations and spreadsheet-based financial analysis.

