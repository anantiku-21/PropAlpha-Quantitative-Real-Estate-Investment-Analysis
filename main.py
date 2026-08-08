import pandas as pd
import numpy as np
import statsmodels.formula.api as smf
import os

def generate_financial_model():
    file_path = 'kc_house_data.csv'
    
    # 1. Structural Check: Ensure data exists
    if not os.path.exists(file_path):
        print(f"CRITICAL ERROR: Dataset '{file_path}' not found in the current directory.")
        return

    print("Ingesting data and running statistical model...")
    raw_df = pd.read_csv(file_path)
    
    # 2. Schema Mapping & Validation
    df = raw_df[['id', 'zipcode', 'price', 'sqft_living', 'bedrooms', 'bathrooms']].copy()
    df.columns = ['property_id', 'zip_code', 'price_usd', 'sq_ft', 'bedrooms', 'bathrooms']
    df = df.dropna()
    df['zip_code'] = df['zip_code'].astype(str)
    
    # 3. Data Engineering: IQR Outlier Removal
    for col in ['price_usd', 'sq_ft', 'bedrooms']:
        Q1 = df[col].quantile(0.25)
        Q3 = df[col].quantile(0.75)
        IQR = Q3 - Q1
        df = df[(df[col] >= Q1 - 1.5 * IQR) & (df[col] <= Q3 + 1.5 * IQR)]
        
    # 4. Statistical Modeling
    model = smf.ols('price_usd ~ sq_ft + bedrooms + bathrooms + C(zip_code)', data=df).fit()
    df['predicted_price'] = model.predict(df)
    
    # 5. Residual Analysis & Target Isolation
    df['residual'] = df['price_usd'] - df['predicted_price']
    
    undervalued = df[df['residual'] < 0].copy()
    # Note: Division for percentage is done natively below, no need to multiply by 100 here
    undervalued['discount_pct'] = np.abs(undervalued['residual']) / undervalued['predicted_price']
    
    top_targets = undervalued.sort_values(by='discount_pct', ascending=False).head(5)
    
    if top_targets.empty:
        print("Model execution complete: Zero statistically undervalued assets found.")
        return

    # ---------------------------------------------------------
    # 6. EXCEL GENERATION & BUSINESS LOGIC (xlsxwriter)
    # ---------------------------------------------------------
    print("Generating dynamic Excel financial model...")
    excel_name = 'PropAlpha_ROI_Model.xlsx'
    writer = pd.ExcelWriter(excel_name, engine='xlsxwriter')
    workbook = writer.book
    
    # Define Professional Formatting
    fmt_currency = workbook.add_format({'num_format': '$#,##0', 'align': 'right'})
    fmt_pct = workbook.add_format({'num_format': '0.00%', 'align': 'right'})
    fmt_header = workbook.add_format({'bold': True, 'bottom': 1, 'bg_color': '#D9D9D9'})
    fmt_input = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1, 'num_format': '0.0%'}) # Yellow for client inputs
    
    # --- SHEET 1: Executive Summary ---
    summary_df = top_targets[['property_id', 'zip_code', 'price_usd', 'predicted_price', 'discount_pct']]
    summary_df.to_excel(writer, sheet_name='Executive Summary', index=False)
    
    ws_summary = writer.sheets['Executive Summary']
    ws_summary.set_column('A:B', 12)
    ws_summary.set_column('C:D', 18, fmt_currency)
    ws_summary.set_column('E:E', 15, fmt_pct)
    for col_num, value in enumerate(summary_df.columns.values):
        ws_summary.write(0, col_num, value, fmt_header)

    # --- SHEET 2: Dynamic ROI Calculator ---
    # Write only the core identifiers; math will be handled by Excel formulas
    roi_df = top_targets[['property_id', 'price_usd', 'predicted_price']]
    roi_df.to_excel(writer, sheet_name='Dynamic ROI Calculator', index=False, startrow=4)
    ws_roi = writer.sheets['Dynamic ROI Calculator']
    
    # Build the Global Assumptions Block (Client-editable)
    ws_roi.write('B1', 'Global Assumptions (Editable)', fmt_header)
    ws_roi.write('B2', 'Down Payment %:')
    ws_roi.write('C2', 0.20, fmt_input) # Default 20%
    ws_roi.write('B3', 'Closing Costs %:')
    ws_roi.write('C3', 0.03, fmt_input) # Default 3%
    
    # Define injected column headers
    calc_headers = ['Down Payment ($)', 'Closing Costs ($)', 'Total Cash Invested', 'Modeled Equity Gain', 'Proj. Cash-on-Cash ROI']
    for col_num, header in enumerate(calc_headers, start=3):
        ws_roi.write(4, col_num, header, fmt_header)

    # Inject dynamic Excel formulas row by row
    for i in range(len(top_targets)):
        row = i + 5 # Data starts on Excel row 6 (0-indexed 5)
        excel_row = row + 1 # For string formula concatenation (e.g., B6)
        
        # Col D: Down Payment ($) = Price * Down Payment %
        ws_roi.write_formula(row, 3, f'=B{excel_row}*$C$2', fmt_currency)
        # Col E: Closing Costs ($) = Price * Closing Costs %
        ws_roi.write_formula(row, 4, f'=B{excel_row}*$C$3', fmt_currency)
        # Col F: Total Cash Invested = Down Payment + Closing Costs
        ws_roi.write_formula(row, 5, f'=D{excel_row}+E{excel_row}', fmt_currency)
        # Col G: Modeled Equity Gain = Predicted Price - Asking Price
        ws_roi.write_formula(row, 6, f'=C{excel_row}-B{excel_row}', fmt_currency)
        # Col H: Cash-on-Cash ROI = Equity Gain / Total Cash Invested
        ws_roi.write_formula(row, 7, f'=G{excel_row}/F{excel_row}', fmt_pct)
        
    ws_roi.set_column('A:A', 12)
    ws_roi.set_column('B:H', 20)
    for col_num, value in enumerate(roi_df.columns.values):
        ws_roi.write(4, col_num, value, fmt_header)

    # --- SHEET 3: Statistical Audit ---
    # Dump the full dataframe so partners can audit the residual math
    df.to_excel(writer, sheet_name='Statistical Audit', index=False)
    ws_audit = writer.sheets['Statistical Audit']
    ws_audit.set_column('A:B', 12)
    ws_audit.set_column('C:C', 15, fmt_currency)
    ws_audit.set_column('D:F', 12)
    ws_audit.set_column('G:H', 15, fmt_currency)
    
    writer.close()
    print(f"SUCCESS: {excel_name} generated in the local directory.")

if __name__ == "__main__":
    generate_financial_model()