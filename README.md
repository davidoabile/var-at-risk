# Inter-Temporal Comparison of Risk Measures: The Effect of Monetary Policy Changes on Value-at-Risk (VaR)

## Overview
This repository contains the analysis and results of a study investigating the impact of monetary policy changes on VaR. The study focuses on how shifts in interest rates, driven by central bank policies, affect the VaR metrics used in risk management and investment strategies. The research employs advanced Monte Carlo simulations to model the potential impacts and provides insights for financial managers and risk intermediaries.

## Hypotheses
- **H0**: Monetary policy changes do not significantly affect the VaR of financial portfolios.
- **H1**: Monetary policy changes significantly affect the VaR of financial portfolios.

## Methodology

### Data Collection
- Historical forex data for AUD/USD and AUD/JPY from the Reserve Bank of Australia (RBA).
- Logarithmic returns calculated for portfolio risk analysis.

### Backtesting
- A framework to examine how portfolios might have reacted to historical monetary policy changes.
- Provides a before-and-after scenario assessment to capture investor sentiment and market reactions.

### Monte Carlo Simulation
- Advanced simulation techniques to model potential future VaR under different cash rate scenarios.
- Incorporates the historical sensitivity of VaR to interest rate changes.

## Key Findings
- Monetary policy changes, such as interest rate adjustments, have a significant impact on VaR.
- Combining VaR with volatility measures offers a more comprehensive risk assessment.
- Short-term volatilities, reflected in a 10-day window, provide valuable insights for dynamic risk management.
- The 99% VaR is more sensitive to interest rate changes, suggesting a conservative approach for risk forecasting.

## Practical Application
- Managers can adjust risk levels based on historical volatility and VaR sensitivity to interest rate changes.
- A combined risk measure (e.g., 95% VaR + volatility) can provide a more realistic assessment of potential losses.
- This approach helps in accounting for unexpected market movements and "fat tail" risks.

## Limitations and Future Work
- Incorporate correlated risk factors and global interest rate movements for a more comprehensive analysis.
- Use advanced distribution models to better account for extreme market behaviors.
- Employ machine learning and AI to measure investor sentiment and improve VaR forecasting.


## Conclusion
This report underscores the critical role of VaR in risk management and investment strategies, highlighting the significant impact of monetary policy changes. By integrating VaR with volatility measures and employing advanced simulations, financial managers can enhance their risk assessment frameworks and better navigate market uncertainties.

## References
- Bodie, Z., Kane, A., & Marcus, A. J. (2001). "Investments." McGraw-Hill.
- Martellini, L., Priaulet, P., & Priaulet, S. (2003). "Fixed-Income Securities: Valuation, Risk Management, and Portfolio Strategies." Wiley Finance.
- Allen, F., & Gale, D. (2004). "Comparative Financial Systems: A Survey." MIT Press.
- Federal Reserve Bank of New York. "Monetary Policy and Financial Conditions: A Cross-Country Study."

## Disclaimer
The code and methods described in this repository are intended for academic and research purposes only. They are not production-ready. If you are interested in developing this code further for practical application, please feel free to collaborate with us.

## License
This project is licensed under the MIT License - see the LICENSE.md file for details.

For more details, please refer to the full report and code in this repository.



