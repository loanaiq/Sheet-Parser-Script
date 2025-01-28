import pandas as pd
import json
import numpy as np

def find_relevant_sheets(file_path, target_keywords):
    """Find sheet names in the workbook that match the target keywords."""
    workbook = pd.ExcelFile(file_path)
    sheet_names = workbook.sheet_names
    print("Sheet Names:", sheet_names)

    index_sheet = pd.read_excel(file_path, sheet_name='Index', header=None)
    index_sheet = index_sheet.iloc[2:].reset_index(drop=True)
    
    matching_sheets = {}
    for keyword in target_keywords:
        matching_rows = index_sheet[index_sheet.iloc[:, 1].str.contains(keyword, case=False, na=False)]
        if not matching_rows.empty:
            sheet_number = matching_rows.iloc[0, 0]
            matching_sheets[keyword] = str(sheet_number)
        else:
            print(f"Keyword '{keyword}' not found in Index sheet.")
    
    return matching_sheets

def parse_table_from_sheet(sheet_data, categories):
    """Extract table data while ignoring rows before/after the table."""
    sheet_data = sheet_data.copy()
    
    # Find the row where 'PARTICULARS' appears in the first column
    particulars_row = None
    for i, row in sheet_data.iterrows():
        if str(row.iloc[0]).strip().upper() == 'PARTICULARS':
            particulars_row = i
            break
    
    if particulars_row is not None:
        # Store the years from the PARTICULARS row
        years_row = sheet_data.iloc[particulars_row]
        # Set the years row as column headers
        sheet_data.columns = years_row
        # Remove the PARTICULARS row and continue processing
        sheet_data = sheet_data.iloc[particulars_row + 1:].reset_index(drop=True)
    else:
        # If PARTICULARS row not found, use original logic
        sheet_data = sheet_data.iloc[2:].reset_index(drop=True)
        sheet_data = sheet_data.dropna(how="all")
        for i, row in sheet_data.iterrows():
            if not row.isnull().all():
                header_index = i
                break
        sheet_data.columns = sheet_data.iloc[header_index]
        sheet_data = sheet_data.iloc[header_index + 1:].reset_index(drop=True)
    
    sheet_data = sheet_data.dropna(how="all")
    
    # Convert only numeric columns to float, keep first column as string
    first_col = sheet_data.iloc[:, 0]
    numeric_cols = sheet_data.iloc[:, 1:]
    
    # Handle the FutureWarning by using infer_objects()
    numeric_cols = numeric_cols.fillna(0).infer_objects(copy=False)
    
    # Try to convert numeric columns to float, replacing errors with 0
    numeric_cols = numeric_cols.apply(pd.to_numeric, errors='coerce').fillna(0)
    
    # Combine back the string column with numeric columns
    result = pd.concat([first_col, numeric_cols], axis=1)
    return result, years_row if particulars_row is not None else None

def numpy_json_encoder(obj):
    """Custom JSON encoder for NumPy types."""
    if isinstance(obj, np.integer):
        return int(obj)
    elif isinstance(obj, np.floating):
        return float(obj)
    elif isinstance(obj, np.ndarray):
        return obj.tolist()
    return obj

def convert_int64_to_int(obj):
    """Convert numeric types to native Python types for JSON serialization."""
    if isinstance(obj, np.integer):
        return int(obj)
    elif isinstance(obj, np.floating):
        return float(obj)
    elif isinstance(obj, np.ndarray):
        return obj.tolist()
    elif isinstance(obj, pd.Series):
        return obj.apply(lambda x: convert_int64_to_int(x) if isinstance(x, (np.number, pd.Series)) else x).tolist()
    elif isinstance(obj, pd.DataFrame):
        return obj.applymap(lambda x: convert_int64_to_int(x) if isinstance(x, np.number) else x).to_dict(orient="records")
    elif isinstance(obj, dict):
        return {key: convert_int64_to_int(value) for key, value in obj.items()}
    elif isinstance(obj, list):
        return [convert_int64_to_int(item) for item in obj]
    return obj

def extract_json_from_sheet(file_path, sheet_number, category_structure):
    """Extract JSON data from a specific sheet."""
    sheet_data = pd.read_excel(file_path, sheet_name=sheet_number, header=None)
    parsed_table, years_row = parse_table_from_sheet(sheet_data, category_structure)
    
    # Extract years from the years_row, excluding the first column (PARTICULARS)
    if years_row is not None:
        years = [str(year) for year in years_row[1:]]
    else:
        years = [str(col) for col in parsed_table.columns[1:]]
    
    json_data = {"years": years}
    
    for category, subcategories in category_structure.items():
        category_data = {}
        for subcategory in subcategories:
            matching_rows = parsed_table[parsed_table.iloc[:, 0].str.strip() == subcategory]
            if not matching_rows.empty:
                row_values = matching_rows.iloc[0, 1:]
                category_data[subcategory] = [convert_int64_to_int(val) for val in row_values]
            else:
                category_data[subcategory] = [0] * len(parsed_table.columns[1:])
        json_data[category] = category_data
    
    return json_data

def main(file_path):
    """Main function to parse the workbook and extract JSON data."""
    target_keywords = ["Balance Sheet", "Profit & Loss", "Ratios"]
    matching_sheets = find_relevant_sheets(file_path, target_keywords)

    if not matching_sheets:
        print("No relevant sheets found.")
        return

    balance_sheet_structure = {
        "SHAREHOLDERS FUND": [
            "Share Capital", "Reserves & Surplus", "Money Received against Warrants", "Networth", 
            "Share Application Money Pending Allotment", "Deffered Government Grants", "Minority Interest"
        ],
        "NON CURRENT LIABILITIES": [
            "Long-term Borrowings", "Secured Long-term Borrowings", "Unsecured Long-term Borrowings (A)+ (B)+ (C) +(D)",
            "Bonds/ Debentures (A)", "Term Loans  (B)", "From banks", "From other parties", 
            "Loans and advances from related parties (C)", "Other Unsecured Long-term Borrowings (D)",
            "Deferred Tax Liabilities", "Other Non Current Liabilities", "Long-term Provisions", "Total Non Current Liabilities"
        ],
        "CURRENT LIABILITIES": [
            "Short-term Borrowings", "Secured Short-term Borrowings", "Unsecured Short-term Borrowings (A)+ (B)+ (C)",
            "Loans repayable on demand  (A)", "From banks", "From other parties", "Loans and advances from related parties (B)",
            "Other Unsecured Short-term Borrowings (C)", "Trade Payables", "Other Current Liabilities", "Short-term Provisions",
            "Total Current Liabilities", "Other Equity & Liabilities", "Total Equity & Liabilities"
        ],
        "NON CURRENT ASSETS": [
            "FIXED ASSET", "Tangible Assets", "Intangible Assets", "Net Block of Assets", "Capital Work in Progress", 
            "Intangible Asset under Development", "Total Fixed Asset", "Non Current Investment", "Deferred Tax Assets (Net)",
            "Long-term Loans & Advances", "Other Non Current Assets", "Total Non Current Assets"
        ],
        "CURRENT ASSETS": [
            "Current Investment", "Inventories", "Trade Receivables", "Cash & Cash Equivalents", 
            "Short-term Loans & Advances", "Other Current Assets", "Total Current Assets", "Other Total Assets", "TOTAL ASSETS"
        ]
    }
    
    profit_loss_structure = {
        "REVENUE": [
            "Revenue from Sale of Products", "Revenue from Sale of Services", "Other Operating Revenues",
            "Gross Sales", "Less:Duties", "Total Revenue from Operations", "Other Income", "Total Revenue"
        ],
        "EXPENSES": [
            "Cost of Materials Consumed", "Purchases of Stock in Trade", "Changes in Inventories of Finished Goods, Work In Progress and Stock In Trade",
            "Total Employee Benefit Expense", "Managerial Remuneration", "Other Employee Benefit Expense", "Total Other Expenses",
            "Payment to Auditors", "Insurance Expenses", "Power and Fuel", "Other Expenses", "EBITDA", "EBITDA %",
            "Finance Costs", "Total Depreciation, Depletion and Amortization Expense", "Total Expenses",
            "Profit before Exceptional and Extraordinary Items and Tax", "Prior Period Items before Tax", "Exceptional Items",
            "Profit before Extraordinary Items and Tax", "Extraordinary Items", "Profit before Tax"
        ],
        "TAX EXPENSE": [
            "Current Tax", "Deferred Tax", "Net Movement in Regulatory Deferral Account Balances related to Profit or Loss and the Related Deferred Tax Movement",
            "Profit/(Loss) for the Period from Continuing Operations", "Profit/(Loss) from Discontinuing Operations",
            "Tax Expense of Discontinuing Operations", "Profit/(Loss) from Discontinuing Operations (After Tax)", "Profit/(Loss)"
        ]
    }
    
    ratios_structure = {
        "PROFITABILITY RATIOS": [
            "Revenue Growth (%)", "EBITDA Margins (%)", "EBT Margins (%)", "PAT Margins (%)", 
            "Return on Equity (%)", "Return on Fixed Assets (%)", "Return on Capital Employed (%)"
        ],
        "LIQUIDITY RATIOS": [
            "Current Ratio", "Quick Ratio"
        ],
        "SOLVENCY RATIOS": [
            "Interest Coverage Ratio", "Long-term Debt/Equity", "Total Assets/Equity", "Total Debt/Equity",
            "Total Debt/Total Assets", "Total Debt/EBITDA"
        ],
        "TURNOVER/EFFICIENCY RATIOS": [
            "Fixed Assets Turnover", "Total Asset Turnover", "Working Capital Turnover", "Inventory Days",
            "Receivables Days", "Payable Days", "Cash Conversion Cycle"
        ],
        "EXPENSES RATIOS": [
            "Raw Material Consumption (% of Sales)", "Total Employee Cost (% of Sales)", "Finance Cost (% of Sales)",
            "Total Other Expenses (% of Sales)"
        ]
    }

    json_data = {}
    for sheet, sheet_number in matching_sheets.items():
        if sheet == "Balance Sheet":
            json_data["balance_sheet"] = extract_json_from_sheet(file_path, sheet_number, balance_sheet_structure)
        elif sheet == "Profit & Loss":
            json_data["profit_loss"] = extract_json_from_sheet(file_path, sheet_number, profit_loss_structure)
        elif sheet == "Ratios":
            json_data["ratios"] = extract_json_from_sheet(file_path, sheet_number, ratios_structure)
    
    with open("parsed_data.json", "w") as json_file:
        json.dump(json_data, json_file, indent=4, default=numpy_json_encoder)

    print("JSON data saved to 'parsed_data.json'")
    return json_data

if __name__ == "__main__":
    file_path = "PERCEPT LIMITED_advance_excel_report.xlsx"  # Replace with your actual file path
    main(file_path)