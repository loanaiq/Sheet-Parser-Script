from fastapi import FastAPI, UploadFile, HTTPException
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import json
import traceback
import numpy as np
from typing import Dict, List, Optional, Union

app = FastAPI()

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload/", 
    response_model=Dict,
    summary="Upload and parse Excel file",
    description="""
    Uploads an Excel file and extracts financial data from Balance Sheet, Profit & Loss, and Ratios sheets.
    Expects an Excel file with an Index sheet that maps sheet numbers to their contents.
    """)
async def receive_excel_file(file: UploadFile):
    """
    Process uploaded Excel file and extract financial data.
    
    Args:
        file (UploadFile): Excel file containing financial statements
        
    Returns:
        dict: Parsed financial data including balance sheet, profit & loss, and ratios
        
    Raises:
        HTTPException: If file processing fails or required sheets are not found
    """
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="File must be an Excel file (.xlsx or .xls)")
    
    try:
        data = await file.read()
        with open("temp_excel.xlsx", "wb") as f:
            f.write(data)
        result = parse('temp_excel.xlsx')
        if not result:
            raise HTTPException(status_code=400, detail="No relevant sheets found in the Excel file")
        return result
    except pd.errors.EmptyDataError:
        raise HTTPException(status_code=400, detail="The Excel file appears to be empty")
    except Exception as e:
        error_detail = f"An error occurred while processing the file: {str(e)}\n{traceback.format_exc()}"
        raise HTTPException(status_code=500, detail=error_detail)

def find_relevant_sheets(file_path: str, target_keywords: List[str]) -> Dict[str, str]:
    """
    Find sheet names in the workbook that match the target keywords.
    
    Args:
        file_path (str): Path to the Excel file
        target_keywords (List[str]): List of keywords to search for in the Index sheet
        
    Returns:
        Dict[str, str]: Mapping of keywords to their corresponding sheet numbers
        
    Raises:
        ValueError: If Index sheet is not found or is improperly formatted
    """
    try:
        workbook = pd.ExcelFile(file_path)
        if 'Index' not in workbook.sheet_names:
            raise ValueError("Index sheet not found in the workbook")

        index_sheet = pd.read_excel(file_path, sheet_name='Index', header=None)
        if index_sheet.empty:
            raise ValueError("Index sheet is empty")

        index_sheet = index_sheet.iloc[2:].reset_index(drop=True)
        
        matching_sheets = {}
        for keyword in target_keywords:
            matching_rows = index_sheet[index_sheet.iloc[:, 1].str.contains(keyword, case=False, na=False)]
            if not matching_rows.empty:
                sheet_number = matching_rows.iloc[0, 0]
                matching_sheets[keyword] = str(sheet_number)
            else:
                print(f"Warning: Keyword '{keyword}' not found in Index sheet.")
        
        return matching_sheets
    except Exception as e:
        raise ValueError(f"Error processing Index sheet: {str(e)}")

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
        return round(float(obj), 2)
    elif isinstance(obj, np.ndarray):
        return [round(float(x), 2) if isinstance(x, (np.floating, float)) else x for x in obj.tolist()]
    elif isinstance(obj, pd.Series):
        return obj.apply(lambda x: round(float(x), 2) if isinstance(x, (np.floating, float)) else 
                       convert_int64_to_int(x) if isinstance(x, (np.number, pd.Series)) else x).tolist()
    elif isinstance(obj, pd.DataFrame):
        return obj.applymap(lambda x: round(float(x), 2) if isinstance(x, (np.floating, float)) else 
                          convert_int64_to_int(x) if isinstance(x, np.number) else x).to_dict(orient="records")
    elif isinstance(obj, dict):
        return {key: convert_int64_to_int(value) for key, value in obj.items()}
    elif isinstance(obj, list):
        return [convert_int64_to_int(item) for item in obj]
    return obj

def extract_json_from_sheet(
    file_path: str, 
    sheet_number: str, 
    category_structure: Dict[str, List[str]]
) -> Dict:
    """
    Extract structured JSON data from a specific sheet.
    
    Args:
        file_path (str): Path to the Excel file
        sheet_number (str): Sheet number to process
        category_structure (Dict[str, List[str]]): Expected structure of categories and subcategories
        
    Returns:
        Dict: Structured financial data with years and categories
        
    Raises:
        ValueError: If sheet data cannot be properly parsed
    """
    try:
        sheet_data = pd.read_excel(file_path, sheet_name=sheet_number, header=None)
        if sheet_data.empty:
            raise ValueError(f"Sheet {sheet_number} is empty")

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
    except Exception as e:
        raise ValueError(f"Error extracting data from sheet {sheet_number}: {str(e)}")

def parse(file_path):
    """Main function to parse the workbook and extract JSON data."""
    target_keywords = ["Balance Sheet", "Profit & Loss", "Ratios", "Profile"]
    matching_sheets = find_relevant_sheets(file_path, target_keywords)

    if not matching_sheets:
        print("No relevant sheets found.")
        return None

    json_data = {}

    # Add profile data first if available
    if "Profile" in matching_sheets:
        try:
            profile_sheet = pd.read_excel(file_path, sheet_name=matching_sheets["Profile"], header=None)
            
            # Find the row index containing "ABOUT THE COMPANY"
            mask = profile_sheet.iloc[:, 0].astype(str).str.contains("ABOUT THE COMPANY", case=False, na=False)
            header_idx = profile_sheet[mask].index
            
            if not header_idx.empty:
                # Get the content from the next row
                content_row = profile_sheet.iloc[header_idx[0] + 1]
                # Get the first non-empty cell from this row
                about_content = next((str(cell) for cell in content_row if pd.notna(cell) and str(cell).strip() != "" and str(cell).lower() != "nan"), None)
                
                if about_content:
                    json_data["metadata"] = {"about_company": about_content.strip()}
                else:
                    json_data["metadata"] = {"about_company": None}
            else:
                json_data["metadata"] = {"about_company": None}
        except Exception as e:
            print(f"Warning: Error processing Profile sheet: {str(e)}")
            json_data["metadata"] = {"about_company": None}

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

    # Process other sheets as before
    for sheet, sheet_number in matching_sheets.items():
        if sheet == "Balance Sheet":
            json_data["balance_sheet"] = extract_json_from_sheet(file_path, sheet_number, balance_sheet_structure)
        elif sheet == "Profit & Loss":
            json_data["profit_loss"] = extract_json_from_sheet(file_path, sheet_number, profit_loss_structure)
        elif sheet == "Ratios":
            json_data["ratios"] = extract_json_from_sheet(file_path, sheet_number, ratios_structure)

    return json_data