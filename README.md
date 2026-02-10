-Project Overview
This project is a comprehensive Excel-based trading performance analysis system designed to evaluate profitability, consistency, and risk across multiple dimensions such as time, strategy, instruments, sessions, phases, and probabilities.
The workbook transforms raw trade-level data into actionable insights using advanced Excel techniques, enabling traders or analysts to identify what works, when it works, and why it works.

-Objectives
Analyze overall and yearly trading performance
Measure profitability using R-based returns
Evaluate strategy effectiveness and strike rates
Identify best-performing time windows, sessions, instruments, and phases
Support data-driven optimization of trading strategies

-Tools & Techniques Used
Microsoft Excel
Advanced Excel Formulas:
SUMIFS
COUNTIFS
IF / IFERROR
VLOOKUP
XLOOKUP

-Data Validation & Named Ranges
Pivot-style aggregations using formulas
Conditional Formatting (performance heatmaps)
Dynamic Charts & Dashboards

-Workbook Structure
Trades Sheet
The core dataset containing trade-level records, including:
Date & Time
Instrument
Strategy type
Trade phase
Session
Risk (R) outcome
Probability classification (High / Medium / Low)

Summary Sheet
A high-level performance dashboard that aggregates data using SUMIFS and COUNTIFS.
Key Metrics:
Total Return (R)
Net Return (R)
Strike Rate (SR)
Total Trades
Wins & Losses
Average Winner & Loser
Profit Factor
Expectancy

-Performance Breakdowns:
Strategy Performance (DSR, WSR, RE, OFS, 4H R, 4H S, RBS)
Instrument Performance (AUDUSD, EURUSD, GBPUSD, XAUUSD, DAX)
Phases Analysis (A, A2, B, C, D, Range)
Days of the Week
Sessions (London, NY, Asia, Low Liquidity)
Probability-Based Returns (High / Medium / Low)
Conditional formatting highlights profitable vs losing segments instantly.

-Yearly Analysis Sheet
A dedicated analytical dashboard focused on year-wise performance.
Includes:
Yearly summary metrics
Monthly Return (R) breakdown
Strategy-wise yearly performance
Instrument-wise contribution
Probability distribution (Pie Chart)
Phase-based performance (Bar Chart)
Dynamic charts provide visual clarity on long-term trends and performance stability.



-Key Calculations & Logic
 Risk-Based Return System
All performance is measured in R (Risk Units) to ensure:
Strategy comparison is normalized
Results are independent of position sizing
Professional risk-adjusted evaluation

-Formula Logic Examples
Return Calculation
=SUMIFS(Trades!$G:$G, Trades!$J:$J, Criteria)

Time-Based Filtering
=SUMIFS(Trades!$J:$J, Trades!$B:$B, ">=12:00", Trades!$B:$B, "<12:30")

Error Handling
=IFERROR(Return / Trades, 0)

-Lookup Usage
VLOOKUP used for structured mappings
XLOOKUP used for flexible and future-proof references

 -Visualizations
Bar charts for phase and strategy performance
Pie charts for probability and instrument contribution
Heatmaps for time-of-day and session performance
Year-over-year comparison charts
All visuals are driven dynamically from formula-based calculations.

-Instrument Performance
EURUSD emerged as the top-performing instrument, contributing approximately +14R to the overall returns.
Despite trading multiple instruments, EURUSD alone accounted for the majority of net profitability, indicating strong alignment between strategy logic and this market’s structure.
Other instruments showed mixed or marginal contributions, reinforcing the importance of instrument selection.

-Probability-Based Performance
Medium probability trades delivered the highest returns, contributing around 31% of total profitable outcomes.
This highlights that moderate-confidence setups with favorable risk-reward ratios outperformed both high- and low-probability trades.
The result confirms that expectancy is driven more by payoff structure than win rate alone.

-Strategy Effectiveness
The OFS (One-Factor Strategy) was the most frequently used strategy and also a major contributor to net returns.
This suggests strong consistency and repeatability in OFS execution across varying market conditions.
Other strategies were used selectively and showed lower impact, making OFS the core performance driver.

-Time-Based Performance
The 19:00–19:30 time window delivered the best entry performance, producing the highest net returns compared to other intraday periods.
This indicates that market structure and liquidity during this window favor the strategy’s entry logic.
Time filtering presents a clear opportunity to improve overall expectancy by focusing on high-performing windows.

-Session Analysis
London Lull recorded a higher win rate than all other sessions, outperforming both London Open and New York sessions.
This challenges the common assumption that high volatility sessions are always optimal.
The data suggests that controlled volatility and cleaner price action during the London Lull align better with the strategy rules.

-Strategic Takeaways
Profitability is concentrated, not evenly distributed.


-Actionable Recommendations
Prioritize EURUSD as the primary trading instrument
Filter trades to medium-probability setups
Make OFS the primary execution strategy
Restrict entries to the 19:00–19:30 window when possible
Increase focus on London Lull session setups
