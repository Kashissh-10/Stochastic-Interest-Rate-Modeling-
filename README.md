**Project Title**: Stochastic Interest Rate Modeling in Excel 

**Objective:** To simulate and analyze short-term interest rate dynamics using two foundational models in quantitative finance — the **Vasicek** and **Cox-Ingersoll-Ross (CIR)** models — and evaluate their performance using historical interest rate data.

**Data Source:** https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202505

**Tools Used:**
- **Microsoft Excel (VBA & Solver ToolPak)**
- **Statistical Techniques:** Maximum Likelihood Estimation (MLE)

**Model Overview:**

**1. Vasicek Model:**
- Assumes mean-reversion with normally distributed innovations.
- Can produce negative interest rates due to its normal distribution assumption.
- Parameters estimated: speed of mean reversion (𝑘), long-term mean (𝜃), and volatility (𝜎).

**2. Cox-Ingersoll-Ross (CIR) Model:**
- Also mean-reverting but ensures non-negativity of rates by using square-root diffusion.
- More realistic for modeling short-term interest rates, especially in low-rate environments.
- Parameters estimated: same as Vasicek.

**Calibration Approach:** Both models were calibrated using **Maximum Likelihood Estimation (MLE)** in Excel.
- Custom VBA functions were created to compute the log-likelihood of observed rate paths under both models.
- Excel Solver was used to optimize the parameters by maximizing the likelihood.

**Key Outcomes & Insights:**

**1. Goodness-of-fit:**
- The CIR model showed better fit to the empirical interest rate data, particularly in capturing rate behavior during periods of low interest rates.
- The Vasicek model, while easier to implement, occasionally predicted negative rates, which is unrealistic in many real-world scenarios.

**2. Volatility Behavior:**
- The CIR model successfully captured the volatility clustering often seen in real-world data due to its dependence on the level of the interest rate.
- Vasicek’s constant volatility assumption proved too simplistic for accurate forecasting in varied market conditions.

**3. Use Cases:**
- CIR model is preferred in regulatory environments and for risk management applications where negative rates are not permissible.
- Vasicek model is simpler and computationally faster — suitable for quick simulations or environments where negative rates might be acceptable (e.g., Eurozone).

**Conclusion:**
This project provided hands-on experience in quantitative finance, model calibration, and time-series analysis using Excel. The CIR model emerged as the more robust and realistic model, especially in post-2008 low-interest regimes. The project also deepened my understanding of stochastic differential equations, calibration techniques, and Excel-based quantitative modeling.
