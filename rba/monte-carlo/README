# Analysis of Monetary Policy Impact on VaR Using Monte Carlo Simulations

## Description

To analyze how changes in monetary policy impact the Value at Risk (VaR), we utilized a structured Monte Carlo simulation technique. This approach involved generating numerous scenarios to model potential portfolio return outcomes under varying cash rate conditions. 

### Formula Used

The formula applied in the simulation for computing simulated VaR is:

### Simulated VaR = Mean VaR + Rate Sensitivity × (Simulated Cash Rate - Mean Cash Rate)

### Key Elements of the Simulation

1. **Statistical Measures**: For each window of data, we computed statistics such as means, medians, standard deviations, variances, portfolio standard deviations, and returns.
2. **Visualization**: To clearly identify the influence of monetary policy, we highlighted rows corresponding to policy announcements in yellow.
3. **Volatility and VaR**: Volatility and VaR calculations were based on historical data for cash rates. These were computed at different confidence levels (90%, 95%, and 99%). The VaR volatility signifies the rate sensitivity in comparison to the standard deviation of cash rates.
4. **Iteration and Portfolio Value**: The simulation ran through 100,000 iterations, considering a portfolio value of $1,000,000.
5. **Data Generation**: Random cash rate values were generated using a normal distribution, the parameters of which (mean and volatility) were derived from historical data.

### Objectives

The primary aim was to quantify the potential financial impact under different scenarios influenced by monetary policy changes. This quantitative analysis helps in understanding the possible fluctuations in portfolio VaR due to external economic decisions, providing a basis for strategic financial planning and risk management.

## Simulation Details

- **Iterations**: 100,000
- **Portfolio Value**: $1,000,000
- **Random Cash Rate Values**: Generated using a normal distribution with parameters derived from historical data.

### VaR Calculation Process

- **Simulated VaR Values**: These were multiplied by the portfolio value ($1,000,000) to convert the VaR from a percentage to a dollar amount.
- **Averages and Standard Deviations**: These metrics were calculated for the simulated VaR values to assess the expected risk and variability under various cash rate scenarios.

## Further Insights

The outcome of the simulations provides valuable insights into how sensitive the portfolio is to changes in interest rates dictated by monetary policy. By understanding these dynamics, institutions can better prepare for potential market conditions that could affect their financial health.

## Adding the Simulation to Excel

To integrate this Monte Carlo simulation into your Excel setup, follow these steps:

1. **Open Excel**: Launch Microsoft Excel on your computer.

2. **Access VBA Editor**:
   - Press `ALT + F11` to open the VBA editor.

3. **Insert a New Module**:
   - In the VBA editor, right-click on any of the items under "VBAProject" (usually named after your workbook).
   - Select `Insert > Module` from the context menu. This creates a new module where you can paste the code.

4. **Copy and Paste the Code**:
   - Copy the provided VBA script.
   - Paste it into the newly created module window in the VBA editor.

5. **Save Your Work**:
   - After pasting, press `CTRL + S` to save your workbook. Make sure to save your workbook as an Excel Macro-Enabled Workbook (*.xlsm) to ensure the macro can be run.

6. **Close the VBA Editor**:
   - You can close the VBA editor and return to your Excel workbook by pressing `ALT + Q`.

7. **Run the Macro**:
   - To execute the macro, go back to Excel and press `ALT + F8` to bring up the list of available macros.
   - Select the `MonteCarloSimulationVaR` macro and click `Run`.

By following these steps, the Monte Carlo simulation will be set up in your Excel environment, allowing you to perform complex risk assessments on your portfolio under various economic scenarios.
