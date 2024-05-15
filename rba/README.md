# Excel VBA Script: Calculate and Format Portfolio Statistics

This repository contains a VBA script designed to calculate and format various statistical measures for a given portfolio. The script processes data for AUD/USD and AUD/JPY currency pairs and computes important financial metrics.


### Description

The script calculates Value at Risk (VaR) over a 10-day window, following the Basel Committee standards. It relies solely on historical data and does not account for economic changes.

For each window, the script computes statistics such as means, medians, standard deviations, variances, portfolio standard deviations, and returns. To visualize the data, rows corresponding to monetary policy announcements are highlighted in yellow. Volatility and VaR are calculated based on historical data for cash rates with different confidence levels (90%, 95%, and 99%). The VaR volatility indicates the rate sensitivity compared to the standard deviation of cash rates.


## Getting Started

Follow these instructions to use the VBA script in Microsoft Excel.

### Prerequisites

- Microsoft Excel

### Steps to Use the Code

1. **Open Excel**: Launch Microsoft Excel on your computer.

2. **Access VBA Editor**:
   - Press `ALT + F11` to open the VBA editor.

3. **Insert a Module**:
   - In the VBA editor, go to `Insert > Module`. A new module window will open.

4. **Copy and Paste the Code**:
   - Copy the provided VBA code below.
   - Paste it into the module window.

5. **Close VBA Editor**:
   - Close the VBA editor by clicking the `X` button or pressing `ALT + Q`.

6. **Prepare Your Data**:
   - Ensure your data is in the following format:
     - Column A: Date
     - Column D: AUD/USD data
     - Column E: AUD/JPY data
     - Column F: Cash Rate
   - Your data should start from row 11.

7. **Run the Macro**:
   - Press `ALT + F8` to open the "Macro" dialog box.
   - Select `CalculateAndFormatStatistics` from the list.
   - Click `Run`.

### Example Output

After running the script, your Excel sheet will display calculated statistics starting from column H and row 10, including:

- Window Number
- Date
- AUD/USD Mean, Median, StdDev, and Variance
- AUD/JPY Mean, Median, StdDev, and Variance
- Portfolio StdDev and Returns
- 90%, 95%, 99% VaR (both percentage and dollar value)
- Cash Rate

### Tips

- Ensure your data starts from row 11 for accurate processing.
- Adjust the `portfolioValue` variable within the code to reflect your actual portfolio value.

### Simulation Details

- **Iterations**: 100,000
- **Portfolio Value**: $1,000,000
- **Random Cash Rate Values**: Generated using a normal distribution with mean and volatility derived from historical data.

### VaR Calculation Process

1. **Simulated VaR Values**: Multiplied by the portfolio value ($1,000,000) to obtain the VaR in dollar terms.
2. **Average and Standard Deviation**: Computed for the simulated VaR values to assess the expected risk and variability under different cash rate scenarios.

### BackTesting Macro

The `BackTesting` macro evaluates the portfolio performance around specified dates (e.g., monetary policy announcements).

# Reserve Bank of Australia - Monetary Policy Decision Dates

We used these dates from the Reserve Bank of Australia for our report:

| Year | Dates of Monetary Policy Decisions               |
|------|-------------------------------------------------|
| 2018 | 7-Feb-18, 7-Mar-18, 4-Apr-18, 2-May-18, 6-Jun-18, 4-Jul-18, 8-Aug-18, 5-Sep-18, 3-Oct-18, 7-Nov-18, 5-Dec-18 |
| 2019 | 5-Feb-19, 5-Mar-19, 2-Apr-19, 7-May-19, 4-Jun-19, 2-Jul-19, 6-Aug-19, 3-Sep-19, 1-Oct-19, 5-Nov-19, 3-Dec-19 |
| 2020 | 4-Feb-20, 3-Mar-20, 19-Mar-20, 7-Apr-20, 5-May-20, 2-Jun-20, 7-Jul-20, 4-Aug-20, 1-Sep-20, 6-Oct-20, 3-Nov-20, 1-Dec-20 |
| 2021 | 2-Feb-21, 2-Mar-21, 6-Apr-21, 4-May-21, 1-Jun-21, 6-Jul-21, 3-Aug-21, 7-Sep-21, 5-Oct-21, 2-Nov-21, 7-Dec-21 |
| 2022 | 1-Feb-22, 1-Mar-22, 5-Apr-22, 3-May-22, 7-Jun-22, 5-Jul-22, 2-Aug-22, 6-Sep-22, 4-Oct-22, 1-Nov-22, 6-Dec-22 |

