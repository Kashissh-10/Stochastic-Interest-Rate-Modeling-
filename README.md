**Project Title**: Stochastic Interest Rate Modeling in Excel 

**Objective:** To simulate and analyze short-term interest rate dynamics using two foundational models in quantitative finance — the **Vasicek** and **Cox-Ingersoll-Ross (CIR)** models — and evaluate their performance using historical interest rate data.

**Data Source:** https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202505

**Tools Used:**
- **Microsoft Excel (VBA & Solver ToolPak)**
- **Statistical Techniques:** Maximum Likelihood Estimation (MLE)

**Technical Implementation:**

**Models Used:**
- **Vasicek Model:** Allows for negative interest rates, has constant volatility.
- **CIR Model:** Enforces non-negativity constraint on interest rates, with volatility dependent on the square root of the rate.

**Calibration Technique:** 
- Used **Maximum Likelihood Estimation (MLE)** to estimate model parameters (speed of mean reversion, long-term mean, and volatility) based on historical interest rate data.
- Built custom MLE optimizations in Excel using Solver and array formulas.

**Monte Carlo Simulations:**
- Simulated **multiple interest rate paths** under each model to observe behavior over time.
- Used **random number generation** and stochastic differential equation discretization.

**Visualization:**
- Plotted the simulated paths, model-implied trajectories, and probability distributions to visually compare results.
- Created **interactive Excel charts** to highlight differences in rate dynamics and volatility behavior.

**Key Outcomes & Insights:**

**1.Model Performance:**
- **CIR Model** provided more **realistic and stable interest rate paths**, especially under low-rate environments, due to its non-negativity feature.
- **Vasicek Model**, while mathematically simpler, occasionally produced **negative rates**, making it less suitable in certain financial applications (e.g., bond pricing).

**2. Volatility Behavior:**
- CIR’s volatility is rate-dependent, which makes it better at capturing increased uncertainty during high-rate regimes.
- Vasicek's constant volatility assumption oversimplified the rate dynamics, particularly in volatile period

**3. Simulation Analysis:**
- CIR simulations showed **mean-reverting and non-negative paths** that aligned better with observed market behavior.
- Vasicek's paths were smoother, but lacked realism during stress periods.

**4. Model Selection:** 
**CIR emerged as the better model** for the given dataset and simulation objectives due to:
- Non-negativity
- Realistic volatility patterns
- Superior fit under MLE calibration (as seen via likelihood metrics and residual analysis)

**Conclusion:**
This project provided hands-on experience in quantitative finance, model calibration, and time-series analysis using Excel. The CIR model emerged as the more robust and realistic model, especially in post-2008 low-interest regimes. The project also deepened my understanding of stochastic differential equations, calibration techniques, and Excel-based quantitative modeling.
