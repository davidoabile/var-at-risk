# Excel VBA Script: Calculate and Format Portfolio Statistics

The script calculates Value at Risk (VaR) over a 10-day window, following the Basel Committee standards. It relies solely on historical data and does not account for economic changes. The script processes data for AUD/USD and AUD/JPY currency pairs and computes important financial metrics.

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
   - Copy the provided VBA code in this folder.
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
