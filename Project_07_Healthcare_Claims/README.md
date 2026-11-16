# Healthcare Insurance Expected Claims Projection | Microsoft Excel

## Project Overview

This project involved developing an Excel-based healthcare insurance claims projection tool for an insurance provider offering medical, dental and vision coverage.

The tool estimates the expected claim amount for an individual policyholder over the duration of their policy. The projection is based on three policyholder characteristics:

- **Age:** Between 20 and 69 years.
- **Gender:** Male or Female.
- **Health Status:** Healthy or Unhealthy.

The policy provides coverage for a maximum of 25 years or until the policyholder reaches age 70, whichever occurs first.

The Excel model uses predefined average claim amounts and age-specific claim probabilities to calculate projected expected claims for each policy year.

## Project Objectives

The main objectives were to:

- Develop a dynamic Excel tool for projecting expected healthcare claims.
- Incorporate differences in average claim amounts between male and female policyholders.
- Apply age-specific claim probabilities.
- Account for the increased claim probability associated with an unhealthy policyholder.
- Automatically adjust projections as the policyholder advances in age.
- Calculate total expected claims over the applicable policy duration.

## Model Assumptions and Parameters

The model uses the following assumptions provided in the project brief:

### 1. Average Claim Amount

| Gender | Average Claim Amount |
|---|---:|
| Male | $687 |
| Female | $492 |

### 2. Probability of Making a Claim by Age Group

| Age Group | Male | Female |
|---|---:|---:|
| 20s | 1.00% | 0.75% |
| 30s | 1.50% | 1.125% |
| 40s | 3.00% | 2.25% |
| 50s | 5.00% | 3.75% |
| 60s | 8.00% | 6.00% |

Female policyholders are assumed to be 25% less likely to make a claim than male policyholders within the same age group.

### 3. Health Status Adjustment

| Health Status | Probability Multiplier |
|---|---:|
| Healthy | 1.00 |
| Unhealthy | 1.50 |

An unhealthy policyholder is assumed to have a 50% higher probability of making a claim than a healthy policyholder.

## Methodology and Calculations

The expected claim amount is calculated using:

**Expected Claim Amount = Average Claim Amount × Adjusted Probability of Making a Claim**

The adjusted probability reflects the policyholder's gender, age group and health status.

For example, the expected claim amount for a female policyholder aged 58 with healthy status is:

- Average claim amount: $492
- Probability for a female in the 50s: 3.75%
- Health status multiplier: 1.00

Expected claim:

$492 × 3.75% × 1.00 = **$18.45**

As the policyholder advances in age, the model automatically identifies the appropriate age group and applies the corresponding probability.

The projection also accounts for the policy eligibility conditions. No claims are projected for policy years in which the policyholder is below age 20 or has reached age 70.

## Excel Implementation

The workbook is organised into two worksheets:

### 1. Projection Tool

This is the main user interface. It contains the policyholder inputs and the projected claims output.

The user specifies:

- Age
- Gender
- Health status

The model then calculates:

- Policy year
- Corresponding age category
- Expected claim amount for each policy year
- Total projected claims over the policy duration

The projection covers up to 25 policy years, subject to the age-70 coverage limit.

### 2. Rate Table

This worksheet contains the assumptions used by the projection model, including:

- Average claim amounts by gender.
- Claim probabilities by age group.
- Health status multipliers.
- Female claim probabilities derived by applying a 25% reduction to the corresponding male probabilities.

Keeping the assumptions in a separate worksheet makes the model easier to review and maintain.

## Excel Functions and Techniques Used

The project applies several Excel functions to automate the calculations:

| Function | Application |
|---|---|
| `IF` | Selects the appropriate average claim amount and health adjustment. |
| `OR` | Checks policy age eligibility conditions. |
| `ROUNDDOWN` | Determines the policyholder's age category for each policy year. |
| `HLOOKUP` | Retrieves the appropriate claim probability from the rate table. |
| `SUM` | Calculates total projected claims across the policy years. |

The formulas link the projection worksheet to the rate table, allowing the output to update when the policyholder's inputs change.

## Key Skills Demonstrated

- Excel-based actuarial modelling
- Expected claims estimation
- Probability-based calculations
- Healthcare insurance risk analysis
- Dynamic formula construction
- Lookup functions and assumption tables
- Policy eligibility modelling
- Financial projection and reporting

## Relevance to Actuarial Work

This project demonstrates the application of probability assumptions and policyholder characteristics in estimating expected insurance claims.

The model reflects important actuarial considerations, including differences in claim experience across demographic groups, adjustments for health status, and coverage eligibility limits.

These techniques are relevant to insurance pricing, expected claims estimation, risk assessment and actuarial modelling.

## Repository Contents

- `Healthcare_Claims.xlsx` – Excel workbook containing the projection tool and supporting rate tables.
