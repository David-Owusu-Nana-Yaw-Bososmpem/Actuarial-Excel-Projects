# RILA Buffer Annuity Pricing and Option Valuation | Excel VBA

## Project Overview

This project involved developing an Excel-based financial modelling tool to calculate the cost of constructing Registered Index-Linked Annuity (RILA), also known as Buffer Annuity, strategies using financial options.

RILAs provide policyholders with exposure to market gains while offering a specified level of protection against market losses. The strategy can be replicated using a combination of a long call option and a short put option.

Using the Black-Scholes option pricing model developed in Project 11 and VBA automation, I built a tool to calculate the option prices for 100 different buffer scenarios and determine their combined cost.

## Project Objectives

- Understand the financial structure of buffer annuities.
- Calculate put strike prices based on buffer percentages.
- Apply the Black-Scholes model to price call and put options.
- Automate option pricing for 100 buffer scenarios using VBA.
- Calculate the net option price for each buffer.
- Summarise the total option cost and average option prices.

## Understanding the RILA Buffer Strategy

A RILA provides protection against an initial percentage of market losses while allowing the policyholder to participate in market upside.

For example, a 10% buffer protects the policyholder against the first 10% of market losses.

If the market declines by 15%, the policyholder experiences a 5% loss after accounting for the buffer.

The strategy in this project uses:

- **Long Call:** Provides exposure to market upside.
- **Short Put:** Helps finance the downside protection provided by the buffer.

The combination is used to replicate the intended RILA payoff structure.

## Methodology and Implementation

### 1. Buffer Strike Price Calculation

Each buffer scenario is defined by three inputs:

- Buffer Strike Price (K)
- Buffer Percentage
- Length of Buffer (Time)

The call option is assumed to be at-the-money, meaning the current underlying asset price equals the call strike price.

The put strike price is calculated using the buffer strike price and buffer percentage.

For example, a buffer strike price of $50 with a 10% buffer produces a put strike price of $45.

### 2. Black-Scholes Option Pricing

The Black-Scholes calculator from Project 11 is used to determine the theoretical prices of the call and put options.

The model uses the following inputs:

- Underlying asset price
- Call strike price
- Put strike price
- Risk-free interest rate
- Volatility
- Time to maturity

The risk-free interest rate and volatility are retrieved using the predefined underlying asset price ranges.

### 3. VBA Automation

VBA is used to automate the valuation of the 100 buffer scenarios.

The automation processes each buffer by:

1. Reading the buffer strike price, percentage and duration.
2. Calculating the corresponding put strike price.
3. Transferring the relevant inputs to the Black-Scholes calculator.
4. Obtaining the call and put option prices.
5. Calculating the net option price.
6. Recording the results in the Buffers worksheet.

This reduces the need to calculate each buffer individually.

### 4. Portfolio-Level Summary

The Total Option Price Tool provides a summary of the calculations across all 100 buffers.

The summary includes:

- Total number of buffers calculated.
- Average call option price.
- Average put option price.
- Total net option price.

These outputs provide a consolidated view of the estimated cost of constructing the buffer strategies.

## Sample Validation

The assignment provides a sample calculation involving four buffer scenarios.

The expected total option price is:

**20.08979402**

This provides a reference for checking the accuracy of the option pricing and aggregation process.

## Workbook Structure

| Worksheet | Description |
|---|---|
| Total Option Price Tool | Summarises the number of buffers, average call and put prices, and total option price. |
| Buffers | Contains the 100 scenarios and their corresponding option pricing calculations. |
| Illustration Tool | Provides the option pricing interface and payoff illustration. |
| Rates | Stores the risk-free interest rate and volatility assumptions. |
| Black Scholes Calculator | Performs the intermediate calculations for call and put option valuation. |

## Tools and Techniques Used

- Microsoft Excel
- Visual Basic for Applications (VBA)
- Black-Scholes option pricing model
- Excel formulas
- VBA loops and automation
- Conditional calculations
- Dynamic cell referencing
- Financial modelling
- Option valuation and aggregation

## Key Skills Demonstrated

- Derivatives pricing
- Financial mathematics
- Structured investment modelling
- RILA and buffer annuity mechanics
- VBA automation
- Option portfolio valuation
- Financial risk management
- Excel-based actuarial modelling

## Actuarial Relevance

Actuaries working in life insurance and financial services may be involved in the valuation, pricing and risk management of investment-linked insurance products.

RILAs require an understanding of financial derivatives, option pricing and the relationship between investment performance and insurance liabilities.

This project demonstrates the practical application of Black-Scholes option valuation and VBA automation to assess the cost of constructing multiple structured investment strategies.

## Repository Contents

- `Buffers_Annuity.xlsm` – Macro-enabled Excel workbook containing the buffer scenarios, Black-Scholes calculator, automated option pricing and portfolio-level summary.
