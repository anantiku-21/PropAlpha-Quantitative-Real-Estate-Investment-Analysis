# PropAlpha-Quantitative-Real-Estate-Investment-Analysis
An automated analytics pipeline that ingests historical real estate transaction data, applies multiple linear regression to isolate mispriced assets, and programmatically generates dynamic Excel financial models for investment analysis.

## Tech Stack
* **Data Engineering:** Python (Pandas, NumPy)
* **Statistical Modeling:** Python (Statsmodels - OLS Regression)
* **Financial Output:** Microsoft Excel (XlsxWriter)

## Pipeline Architecture

### 1. Data Architecture & Engineering
Ingests raw `kc_house_data.csv` transaction datasets and validates schema mapping. Applies strict Interquartile Range (IQR) filtering across pricing and square footage to remove extreme outliers, mitigating market variance and stabilizing the regression model.

### 2. Predictive Statistical Modeling
Fits an Ordinary Least Squares (OLS) regression model to calculate theoretical baseline asset valuations. 
* Utilizes categorical encoding for spatial data (ZIP codes) to prevent invalid ordinal math.
* Calculates the residual error term between actual asking price and modeled market value. 
* Properties exhibiting a severe negative residual are mathematically isolated as undervalued.

### 3. Complex Financial Modeling
Bypasses static dashboards to automatically compile a multi-sheet Excel deliverable. The Python script uses `xlsxwriter` to inject live formulas directly into the workbook, translating raw statistical anomalies into a dynamic ROI calculator where stakeholders can adjust closing costs and down payment percentages.

## Execution Requirements
1. Install dependencies:
   ```bash
   pip install pandas numpy statsmodels xlsxwriter
