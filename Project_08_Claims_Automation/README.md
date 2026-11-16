# Automated Healthcare Insurance Claims Projection for 100 Policyholders | Excel VBA

## Project Overview

This project extends the healthcare insurance expected claims projection model developed in Project 7.

An insurance provider supplied a dataset containing 100 policyholders, with information on their gender, age and health condition. The objective was to develop an Excel VBA solution that calculates the total projected healthcare claims for each policyholder over the applicable duration of their insurance policy.

Rather than manually calculating expected claims for each individual, I used VBA automation to apply the existing actuarial projection methodology across the policyholder dataset.

The solution combines probability-based expected claims calculations with VBA programming to support efficient processing of multiple policyholders.

## Project Objectives

The main objectives were to:

- Automate the calculation of projected claims for 100 policyholders using VBA.
- Apply the healthcare insurance assumptions established in Project 3.
- Calculate expected claims across each policyholder's eligible policy years.
- Aggregate annual expected claims into a single total projected claims amount.
- Display the calculated total beside each policyholder's existing information.
- Reduce repetitive manual calculations through VBA programming.

## Dataset Description

The dataset contains 100 policyholders and the following information:

| Variable | Description |
|---|---|
| Name | Policyholder identifier |
| Gender | Male or Female |
| Age | Policyholder's current age |
| Condition | Health status represented numerically |
| Total Projected Claims | Calculated output |

The health condition variable is coded as follows:

| Code | Health Condition |
|---|---|
| 1 | Unhealthy |
| 2 | Healthy |

## Actuarial Assumptions

The model retains the assumptions from Project 3.

### 1. Average Claim Amount

| Gender | Average Claim Amount |
|---|---:|
| Male | $687 |
| Female | $492 |

### 2. Probability of Making a Claim

| Age Group | Male | Female |
|---|---:|---:|
| 20s | 1.00% | 0.75% |
| 30s | 1.50% | 1.125% |
| 40s | 3.00% | 2.25% |
| 50s | 5.00% | 3.75% |
| 60s | 8.00% | 6.00% |

Female policyholders are assumed to be 25% less likely to make a claim than male policyholders in the same age group.

### 3. Health Status Adjustment

| Health Status | Probability Multiplier |
|---|---:|
| Healthy | 1.00 |
| Unhealthy | 1.50 |

An unhealthy policyholder has a 50% higher probability of making a claim than a healthy policyholder.

### 4. Policy Duration and Eligibility

- The policy is available to individuals aged 20 to 69.
- Coverage lasts for a maximum of 25 years or until the policyholder reaches age 70, whichever occurs first.
- Each policyholder is assumed to purchase the policy on their birthday.
- The policyholder advances one year in age at the beginning of each subsequent policy year.

## Methodology and Calculations

The model uses the expected claims calculation established in Project 3:

**Annual Expected Claim = Average Claim Amount × Probability of Making a Claim**

The probability is adjusted according to the policyholder's age group, gender and health status.

For each policyholder, the model calculates the expected claim amount for every eligible policy year and adds the annual amounts to obtain the total projected claims.

**Total Projected Claims = Sum of Expected Claims Across All Eligible Policy Years**

For example, a policyholder aged 46 will have their expected claims calculated using the relevant age-based probabilities as they progress through the policy years. When the policyholder enters a new age bracket, the corresponding claim probability is applied.

The projection stops when the policyholder reaches age 70 or completes the maximum 25-year policy duration.

## VBA Automation

The main programming component of this project is the use of VBA to automate the claims projection process across the 100 policyholders.

The workbook includes:

- A `Policy_Data` worksheet containing the 100 policyholder records.
- A `Projection Tool` worksheet containing the individual claims projection model.
- A `Rate Table` worksheet containing the actuarial assumptions.
- Embedded VBA procedures, including a looping routine and a procedure named `proj_claim`.

The VBA implementation is designed to process the policyholder records, apply the individual claims projection calculations and place the resulting total projected claims in the fifth column of the dataset.

Using a loop avoids the need to repeat the calculation manually for each policyholder.

## Workbook Structure

The Excel workbook contains three main worksheets:

### 1. Projection Tool

Contains the individual policyholder claims projection model developed in Project 3. It calculates annual expected claims and aggregates them over the eligible policy duration.

### 2. Rate Table

Contains the actuarial assumptions used in the calculations, including average claim amounts, age-specific claim probabilities and health status multipliers.

### 3. Policy_Data

Contains the 100 policyholder records, including their names, gender, age and health condition.

The fifth column, `Total Projected Claims`, is designated for the VBA-generated output.

## Excel and Programming Skills Demonstrated

- Visual Basic for Applications (VBA)
- VBA loops and conditional statements
- Excel worksheet and cell manipulation
- Automation of repetitive calculations
- Expected claims modelling
- Probability-based calculations
- Data processing and aggregation
- Integration of Excel formulas with VBA procedures

## Relevance to Actuarial Work

This project demonstrates how actuarial calculations can be automated and applied to a portfolio of insurance policyholders.

It connects individual expected claims projections with portfolio-level data processing, illustrating how probability assumptions and policyholder characteristics can be used to estimate future claims.

The skills demonstrated are relevant to:

- Actuarial pricing and expected claims estimation
- Insurance portfolio analysis
- Claims modelling
- Risk assessment
- Actuarial data processing and reporting
- Automation of recurring actuarial calculations

## Repository Contents

- `Claims_Automation.xlsm` – Excel macro-enabled workbook containing the policyholder dataset, projection model, rate table and VBA implementation.

experience or a complete commercial pricing model.
