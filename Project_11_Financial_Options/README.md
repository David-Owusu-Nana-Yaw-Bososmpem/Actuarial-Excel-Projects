# Black-Scholes Option Pricing Calculator | Microsoft Excel

## Project Overview

This project involved developing an Excel-based option pricing calculator using the Black-Scholes model to determine the theoretical prices of European call and put options.

The calculator allows users to select an option type and provide the underlying asset price, strike price and time to maturity. Risk-free interest rates and volatility are automatically assigned based on the underlying asset price using predefined assumptions.

The model also includes a payoff diagram to illustrate how the option's payoff changes across different underlying asset prices.

## Project Objectives

- Develop an Excel calculator for pricing call and put options.
- Implement the Black-Scholes option pricing model.
- Automate the selection of risk-free interest rates and volatility.
- Calculate the intermediate values required for option pricing.
- Create a dynamic payoff diagram to illustrate option behaviour.
- Apply Excel functions to financial and derivative valuation.

## Model Assumptions

The calculator uses the following assumptions:

- The underlying asset does not pay dividends.
- The option is valued using the Black-Scholes model.
- Time to maturity is measured in years.
- Risk-free interest rates and volatility are determined using predefined underlying asset price ranges.

### Risk-Free Interest Rate and Volatility Assumptions

| Underlying Asset Price | Risk-Free Rate | Volatility |
|---|---:|---:|
| 0–less than 25 | 2% | 50% |
| 25–less than 50 | 3% | 40% |
| 50–less than 100 | 4% | 25% |
| 100–less than 200 | 5% | 15% |
| 200 and above | 6% | 10% |

These assumptions are stored in a separate worksheet and automatically retrieved based on the underlying asset price.

## Methodology and Implementation

### 1. Option Pricing Calculator

The main calculator accepts the following inputs:

| Input | Description |
|---|---|
| Option Type | Call or Put |
| S(0) | Current price of the underlying asset |
| K | Strike price |
| T | Time to maturity in years |
| r | Risk-free interest rate |
| σ | Volatility |

The calculator automatically determines the applicable risk-free interest rate and volatility using the predefined price ranges.

### 2. Black-Scholes Model

The Black-Scholes model is used to calculate the theoretical prices of European call and put options.

The model incorporates the underlying asset price, strike price, time to maturity, risk-free interest rate and volatility.

The calculator first determines the intermediate values \(d_1\) and \(d_2\), which are then used with the standard normal cumulative distribution function to calculate the option prices.

The final output depends on the option type selected by the user.

### 3. Payoff Diagram

A payoff diagram is incorporated into the calculator to illustrate how the option's payoff changes at different underlying asset prices.

The chart uses a range of underlying asset prices generated from an adjustable starting value and increment.

This provides a visual representation of the relationship between the underlying asset price and the option's value at expiration.

## Sample Validation

The calculator was tested against the sample scenarios provided in the assignment guide.

| Option Type | S(0) | K | Time to Maturity | Risk-Free Rate | Volatility | Option Price |
|---|---:|---:|---:|---:|---:|---:|
| Call | 20 | 25 | 1 year | 2% | 50% | 2.489 |
| Put | 33 | 33 | 2 years | 3% | 40% | 6.213 |
| Put | 50 | 60 | 400 days | 4% | 25% | 10.065 |
| Call | 225 | 220 | 60 days | 6% | 10% | 8.243 |

The calculated call and put prices were compared against the expected results provided in the project guide.

## Workbook Structure

| Worksheet | Description |
|---|---|
| Illustration Tool | Main user interface for entering inputs, displaying option prices and viewing the payoff diagram. |
| Rates | Contains the underlying asset price ranges and corresponding risk-free rates and volatility assumptions. |
| Black Scholes Calculator | Performs the intermediate calculations and determines call and put option prices. |

## Excel Functions and Techniques Used

- `IF` – Determines the appropriate option price and applies conditional logic.
- `AND` – Evaluates underlying asset price ranges.
- `LN` – Calculates natural logarithms.
- `EXP` – Calculates exponential values.
- `SQRT` – Calculates square roots.
- `NORM.S.DIST` – Returns standard normal cumulative distribution values.
- Dynamic cell referencing – Connects the calculator, assumptions and pricing worksheets.
- Excel charts – Visualises option payoff behaviour.

## Key Skills Demonstrated

- Derivatives pricing
- Black-Scholes option valuation
- Financial mathematics
- Probability distributions
- Excel financial modelling
- Automated calculations
- Risk and volatility analysis
- Data visualisation

## Actuarial Relevance

Financial derivatives are important instruments in financial risk management, particularly when managing exposure to market movements.

Actuaries working in insurance and financial services may use options and other derivatives as part of investment analysis, asset-liability management and hedging strategies.

This project demonstrates the practical application of financial mathematics, probability distributions and Excel modelling to derivative valuation.

## Repository Contents

- `Financial_Options.xlsx` – Excel workbook containing the Black-Scholes calculator, underlying asset price assumptions and option payoff diagram.

