# Football Results Data Analysis (Excel Report Generator)

A Python project that takes raw **football (soccer) match results** in CSV format and produces a **styled Excel workbook** with insightful analysis — _no Pandas required_. This was started as a personal project and will continue to evolve over time with new features and smarter insights.

---

## Project Overview
This script processes match results data from a CSV (`results.csv`) and generates a single Excel report (`football_analysis_report.xlsx`) containing:

1. **Raw Data Sheet**  
   - All results loaded from CSV  
   - Improved readability with alternating row colors  
   - Auto-sized columns  

2. **Yearly Analysis Sheet**  
   - Aggregated statistics for each season:  
     - Total Goals scored  
     - Number of Home Wins  
     - Number of Away Wins  
     - Win Difference (Home − Away)  
   - Visualizations:  
     - Bar Chart → Goals per Season  
     - Line Chart → Home vs Away Win Difference  

3. **Team Win Rates Sheet**  
   - Game counts and win rates per team (home vs away split)  
   - Columns include:  
     - Home/Away Games Played  
     - Home/Away Wins  
     - Home/Away Win Rates (%)  

---

## How to Run

### 1. Install dependencies  
```bash
pip install openpyxlgrouping which then completed the third sheet for me.
```
### 2. Prepare your data
Place a results.csv file in the same folder as the script. The CSV must include at least:
Season, HomeTeam, AwayTeam, FTHG, FTAG, FTR
``` check the one under the excel folder ```

Season → e.g. 2018/19
HomeTeam → Home team name
AwayTeam → Away team name
FTHG → Full-Time Home Goals (number)
FTAG → Full-Time Away Goals (number)
FTR → Result (H = Home Win, A = Away Win, D = Draw)

### 3. Run the script
```python football_report.py```

football_analysis_report.xlsx

## Example Output

* Raw Football Results – all match data, styled
* Yearly Analysis – aggregated stats + charts
* Team Win Rates – per team performance split into home/away

## Features

* Works without pandas or heavy libraries
* Consistent styled Excel sheets (headers colored, zebra-striped rows)
* Auto-sized columns for readability
* Inline bar and line charts in Excel
* Modular functions for easy extension

## Future Improvements

* Add draw rate calculations
* Split seasonal goals (home vs away totals)
* Add trend insights (e.g. changes in home advantage)
* More chart types (stacked bars, win %, etc.)
* Turn project into a CLI tool or small web app

## Author & Motivation

Fiston Kilele
A personal project to practice Python and explore football results data. Originally a utility script, now being improved step by step into a polished analysis/reporting tool.
