import pandas as pd
import os
from datetime import datetime
import numpy as np

def find_excel_files():
    """
    Scan current directory and all subfolders for Excel files containing
    rpgBorrowing and rpgRedemptions sheets.
    """
    excel_files = []
    current_dir = os.getcwd()
    
    print("Scanning directories for Excel files...")
    
    # Walk through all directories and subdirectories
    for root, dirs, files in os.walk(current_dir):
        for file in files:
            if file.endswith('.xlsx') and not file.startswith('~'):  # Exclude temp files
                file_path = os.path.join(root, file)
                
                # Check if file contains required sheets
                try:
                    excel_file = pd.ExcelFile(file_path)
                    sheet_names = excel_file.sheet_names
                    
                    if 'rpgBorrowing' in sheet_names and 'rpgRedemptions' in sheet_names:
                        excel_files.append(file_path)
                        rel_path = os.path.relpath(file_path, current_dir)
                        print(f"✓ Found valid file: {rel_path}")
                    
                    excel_file.close()
                    
                except Exception as e:
                    # Skip files that can't be read
                    continue
    
    return excel_files

def format_number(value):
    """Format numbers with commas as thousand separators"""
    if pd.isna(value) or value is None:
        return None
    
    try:
        # Convert to float first
        num_value = float(value)
        
        # Check if it's a whole number
        if num_value == int(num_value):
            return int(num_value)
        else:
            return num_value
            
    except (ValueError, TypeError):
        return value

def extract_gfb_data(source_file):
    """
    Extract and map data from Excel file (rpgBorrowing and rpgRedemptions sheets)
    to create standardized GFB_DATA output with hardcoded headers.
    """
    
    print(f"Processing file: {source_file}")
    
    # Hardcoded headers for all 108 columns (54 BORR + 54 REDEM)
    BORROWING_CODES = [
        "GFB.BORR.FINBUDGET.LOANFIN.M",
        "GFB.BORR.FINBUDGET.M", 
        "GFB.BORR.BREAKPURP.M",
        "GFB.BORR.FEDBUDGET.M",
        "GFB.BORR.FINMARKETFUND.M",
        "GFB.BORR.FMSLOANEXP.M",
        "GFB.BORR.FMSLOANWINDAGE.M",
        "GFB.BORR.INVREDFUND.M",
        "GFB.BORR.ECONSTABFUND.M",
        "GFB.BORR.ESFLOANRECAPMEAS.M",
        "GFB.BORR.ESFLOANKFW23.M",
        "GFB.BORR.ESFLOANENERGCRIS.M",
        "GFB.BORR.ESFLOANKFW26.M",
        "GFB.BORR.SPECFUNDBUND.M",
        "GFB.BORR.RESTRFUND.M",
        "GFB.BORR.HARDCOALEQFUND.M",
        "GFB.BORR.FEDRAILFUND.M",
        "GFB.BORR.COMPFUND.M",
        "GFB.BORR.REDFUNDLIAB.M",
        "GFB.BORR.ERPSPECFUND.M",
        "GFB.BORR.GERUNITFUND.M",
        "GFB.BORR.EQUALBURFUND.M",
        "GFB.BORR.BREAKDEPTTYPE.M",
        "GFB.BORR.BREAKDEPTTYPE.FEDSEC.M",
        "GFB.BORR.BREAKDEPTTYPE.CONVFEDSEC.M",
        "GFB.BORR.BREAKDEPTTYPE.FEDBOND.M",
        "GFB.BORR.BREAKDEPTTYPE.30YFEDBOND.M",
        "GFB.BORR.BREAKDEPTTYPE.15YFEDBOND.M",
        "GFB.BORR.BREAKDEPTTYPE.10YFEDBOND.M",
        "GFB.BORR.BREAKDEPTTYPE.7YFEDBOND.M",
        "GFB.BORR.BREAKDEPTTYPE.FEDNOTE.M",
        "GFB.BORR.BREAKDEPTTYPE.FEDTREASNOTE.M",
        "GFB.BORR.BREAKDEPTTYPE.TREASDISCPAP.M",
        "GFB.BORR.BREAKDEPTTYPE.INFFEDSEC.M",
        "GFB.BORR.BREAKDEPTTYPE.GREENFEDSEC.M",
        "GFB.BORR.BREAKDEPTTYPE.SUPSECFEDGOV.M",
        "GFB.BORR.BREAKDEPTTYPE.LOANSUPFEDGOV.M",
        "GFB.BORR.BREAKDEPTTYPE.ESFINVFEDGOVSEC.M",
        "GFB.BORR.BREAKDEPTTYPE.OTHERFEDSEC.M",
        "GFB.BORR.BREAKDEPTTYPE.PROMNOTE.M",
        "GFB.BORR.BREAKDEPTTYPE.OTHERLOANORDDEPT.M",
        "GFB.BORR.OWNHOLD.M",
        "GFB.BORR.OWNHOLD.CONVFEDSEC.M",
        "GFB.BORR.OWNHOLD.30YFEDBOND.M",
        "GFB.BORR.OWNHOLD.15YFEDBOND.M",
        "GFB.BORR.OWNHOLD.10YFEDBOND.M",
        "GFB.BORR.OWNHOLD.7YFEDBOND.M",
        "GFB.BORR.OWNHOLD.FEDNOTE.M",
        "GFB.BORR.OWNHOLD.FEDTREASNOTE.M",
        "GFB.BORR.OWNHOLD.TREASDISCPAP.M",
        "GFB.BORR.OWNHOLD.INFFEDSEC.M",
        "GFB.BORR.OWNHOLD.GREENFEDSEC.M",
        "GFB.BORR.OWNHOLD.OTHERFEDSEC.M",
        "GFB.BORR.ESFINVFEDGOVSEC.M"
    ]
    
    BORROWING_DESCRIPTIONS = [
        " Financing Federal budget, special funds; loan financing FMS & ESF",
        "Gross Borrowing Requirement: Financing Federal budget and special funds",
        "Gross Borrowing Requirement: Breakdown by purpose",
        "Gross Borrowing Requirement: Federal budget",
        "Gross Borrowing Requirement: Financial Market Stabilisation Fund",
        "Gross Borrowing Requirement: FMS loans for expenses acc. to section 9 (1)  StFG",
        "Gross Borrowing Requirement: FMS loans for wind-up agencies acc. to section 9 (5)  StFG",
        "Gross Borrowing Requirement: Investment and Redemption Fund",
        "Gross Borrowing Requirement: Economic Stabilisation Fund",
        "Gross Borrowing Requirement: ESF loans for recapitalisation measures acc. to section 22 StFG",
        "Gross Borrowing Requirement: ESF loans for KfW acc. to section 23 StFG",
        "Gross Borrowing Requirement: ESF loans to mitigate consequences of the energy crisis acc. to sect. 26a (1) no. 1-4 StFG",
        "Gross Borrowing Requirement: ESF loans to KfW acc. to section 26a (1) no 5 StFG",
        "Gross Borrowing Requirement: Special Fund for the Bundeswehr",
        "Gross Borrowing Requirement: Restructuring Fund",
        "Gross Borrowing Requirement: Hard Coal Equalisation Fund",
        "Gross Borrowing Requirement: Federal Railways Fund",
        "Gross Borrowing Requirement: Compensation Fund",
        "Gross Borrowing Requirement: Redemption Fund for Inherited Liabilities",
        "Gross Borrowing Requirement: ERP Special Fund",
        "Gross Borrowing Requirement: German Unity Fund",
        "Gross Borrowing Requirement: Equalisation of Burdens Fund",
        "Gross Borrowing Requirement: Breakdown by debt type",
        "Gross Borrowing Requirement: Breakdown by debt type: Federal securities",
        "Gross Borrowing Requirement: Breakdown by debt type: Conventional Federal securities",
        "Gross Borrowing Requirement: Breakdown by debt type: Federal bonds",
        "Gross Borrowing Requirement: Breakdown by debt type: 30-year Federal bonds",
        "Gross Borrowing Requirement: Breakdown by debt type: 15-year Federal bonds",
        "Gross Borrowing Requirement: Breakdown by debt type: 10-year Federal bonds",
        "Gross Borrowing Requirement: Breakdown by debt type: 7-year Federal bonds",
        "Gross Borrowing Requirement: Breakdown by debt type: Federal notes",
        "Gross Borrowing Requirement: Breakdown by debt type: Federal Treasury notes",
        "Gross Borrowing Requirement: Breakdown by debt type: Treasury discount paper",
        "Gross Borrowing Requirement: Breakdown by debt type: Inflation-linked Federal securities",
        "Gross Borrowing Requirement: Breakdown by debt type: Green Federal securities",
        "Gross Borrowing Requirement: Breakdown by debt type: Supplementary securities issued by the federal government",
        "Gross Borrowing Requirement: Breakdown by debt type: Loans from suppl. issues by the federal government for the ESF acc. to sect. 26b StFG",
        "Gross Borrowing Requirement: Breakdown by debt type: ESF investment in federal government securities acc. to sect. 26b (5) StFG",
        "Gross Borrowing Requirement: Breakdown by debt type: Other Federal securities",
        "Gross Borrowing Requirement: Breakdown by debt type: Promissory notes",
        "Gross Borrowing Requirement: Breakdown by debt type: Other loans and ordinary debts",
        "Gross Borrowing Requirement: Own holdings",
        "Gross Borrowing Requirement: Own holdings: Conventional Federal securities",
        "Gross Borrowing Requirement: Own holdings: 30-year Federal bonds",
        "Gross Borrowing Requirement: Own holdings: 15-year Federal bonds",
        "Gross Borrowing Requirement: Own holdings: 10-year Federal bonds",
        "Gross Borrowing Requirement: Own holdings: 7-year Federal bonds",
        "Gross Borrowing Requirement: Own holdings: Federal notes",
        "Gross Borrowing Requirement: Own holdings: Federal Treasury notes",
        "Gross Borrowing Requirement: Own holdings: Treasury discount paper",
        "Gross Borrowing Requirement: Own holdings: Inflation-linked Federal securities",
        "Gross Borrowing Requirement: Own holdings: Green Federal securities",
        "Gross Borrowing Requirement: Own holdings: Other Federal securities",
        "Gross Borrowing Requirement: ESF investment in federal government securities acc. to sect. 26b (5) StFG"
    ]
    
    # Create redemption codes and descriptions by replacing BORR with REDEM and updating descriptions
    REDEMPTION_CODES = [code.replace('BORR', 'REDEM') for code in BORROWING_CODES]
    REDEMPTION_DESCRIPTIONS = [desc.replace('Gross Borrowing Requirement:', 'Redemption Payments:').replace(' Financing Federal budget', 'Redemption Payments: Financing Federal budget') for desc in BORROWING_DESCRIPTIONS]
    
    # Combine all headers
    ALL_CODES = [None] + BORROWING_CODES + REDEMPTION_CODES  # None for date column
    ALL_DESCRIPTIONS = [None] + BORROWING_DESCRIPTIONS + REDEMPTION_DESCRIPTIONS
    
    try:
        # Read the source sheets
        print("Reading rpgBorrowing sheet...")
        borrowing_df = pd.read_excel(source_file, sheet_name='rpgBorrowing', header=None)
        
        print("Reading rpgRedemptions sheet...")
        redemptions_df = pd.read_excel(source_file, sheet_name='rpgRedemptions', header=None)
        
        # Extract dates from row 14 (index 13), columns B+ (index 1+)
        print("Extracting dates...")
        date_row = borrowing_df.iloc[13, 1:].dropna()
        
        # Convert dates to YYYY-MM format
        dates = []
        for date_val in date_row:
            if pd.notna(date_val):
                if isinstance(date_val, datetime):
                    dates.append(date_val.strftime('%Y-%m'))
                else:
                    # Try to parse string dates
                    try:
                        parsed_date = pd.to_datetime(date_val)
                        dates.append(parsed_date.strftime('%Y-%m'))
                    except:
                        dates.append(str(date_val))
        
        print(f"Found {len(dates)} date columns")
        
        # Extract borrowing data with proper row mapping
        print("Extracting borrowing data...")
        borrowing_data = []
        
        # Define proper row mapping for 54 measures (based on actual Excel structure)
        borrowing_row_mapping = [
            # Columns 1-22: Basic measures (rows 15-36, indices 14-35)
            14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35,
            # Columns 23-41: Breakdown by debt type (rows 37-55, indices 36-54)  
            36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54,
            # Columns 42-54: Own holdings (rows 90-102, indices 89-101)
            89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101
        ]
        
        for i, row_idx in enumerate(borrowing_row_mapping):
            if row_idx < len(borrowing_df):
                row_data = borrowing_df.iloc[row_idx, 1:len(dates)+1].values
                borrowing_data.append(row_data)
                if i >= 41:  # Own holdings section
                    desc = borrowing_df.iloc[row_idx, 0] if row_idx < len(borrowing_df) else "N/A"
                    print(f"  Col {i+1}: Row {row_idx+1} -> '{desc}' = {row_data[0] if len(row_data) > 0 else 'N/A'}")
            else:
                # Pad with NaN if row doesn't exist
                borrowing_data.append([np.nan] * len(dates))
        
        # Extract redemption data with proper row mapping
        print("Extracting redemption data...")
        redemption_data = []
        
        # Use same row mapping for redemptions
        redemption_row_mapping = borrowing_row_mapping  # Same structure
        
        for i, row_idx in enumerate(redemption_row_mapping):
            if row_idx < len(redemptions_df):
                row_data = redemptions_df.iloc[row_idx, 1:len(dates)+1].values
                redemption_data.append(row_data)
                if i >= 41:  # Own holdings section
                    desc = redemptions_df.iloc[row_idx, 0] if row_idx < len(redemptions_df) else "N/A"
                    print(f"  Col {i+55}: Row {row_idx+1} -> '{desc}' = {row_data[0] if len(row_data) > 0 else 'N/A'}")
            else:
                # Pad with NaN if row doesn't exist
                redemption_data.append([np.nan] * len(dates))
        
        # Combine all data
        all_data = borrowing_data + redemption_data
        
        # Create the final DataFrame
        print("Creating final output...")
        
        # Create data rows: each column becomes a row, with date as first column
        output_data = []
        
        # Add header rows
        output_data.append(ALL_CODES)
        output_data.append(ALL_DESCRIPTIONS)
        
        # Add data rows
        for i, date in enumerate(dates):
            row = [date]  # Start with date
            
            # Add data for each measure (borrowing + redemption)
            for measure_idx in range(len(all_data)):
                if i < len(all_data[measure_idx]):
                    value = all_data[measure_idx][i]
                    # Convert to numeric, keep NaN for missing values
                    if pd.isna(value):
                        row.append(None)
                    else:
                        try:
                            row.append(float(value))
                        except:
                            row.append(value)
                else:
                    row.append(None)
            
            output_data.append(row)
        
        # Create DataFrame and format numbers
        print("Formatting numbers with thousand separators...")
        formatted_data = []
        for row in output_data:
            formatted_row = []
            for i, cell in enumerate(row):
                if i == 0 or i < 2:  # Date column or header rows
                    formatted_row.append(cell)
                else:
                    formatted_row.append(format_number(cell))
            formatted_data.append(formatted_row)
        
        output_df = pd.DataFrame(formatted_data)
        
        # Generate output filename with timestamp and source file name
        timestamp = datetime.now().strftime("%Y%m%d")
        source_basename = os.path.splitext(os.path.basename(source_file))[0]
        output_filename = f"GFB_DATA_HARDCODED_{source_basename}_{timestamp}.xlsx"
        
        # Save to Excel with number formatting
        print(f"Saving to {output_filename}...")
        
        # Create Excel writer with formatting options
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            output_df.to_excel(writer, sheet_name='DATA', index=False, header=False)
            
            # Get worksheet and apply number formatting
            worksheet = writer.sheets['DATA']
            
            # Apply thousand separator formatting to data cells (skip headers)
            for row_num in range(3, len(formatted_data) + 1):  # Start from row 3 (after headers)
                for col_num in range(2, len(ALL_CODES) + 1):  # Start from column B (after dates)
                    cell = worksheet.cell(row=row_num, column=col_num)
                    if cell.value is not None and isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0'
        
        print(f"Successfully created {output_filename}")
        print(f"   - {len(formatted_data)-2} data rows")
        print(f"   - {len(ALL_CODES)-1} measure columns (54 BORR + 54 REDEM)")
        print(f"   - Date range: {dates[0] if dates else 'N/A'} to {dates[-1] if dates else 'N/A'}")
        print(f"   - Numbers formatted with thousand separators")
        
        return output_filename
        
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        raise

if __name__ == "__main__":
    print("=== GFB Data Extraction Script ===")
    print("Scanning current directory and subfolders for Excel files...")
    
    try:
        # Find all valid Excel files
        excel_files = find_excel_files()
        
        if not excel_files:
            print("No valid Excel files found!")
            print("Looking for files with 'rpgBorrowing' and 'rpgRedemptions' sheets")
            exit(1)
        
        print(f"\nFound {len(excel_files)} valid Excel file(s)")
        
        # Process each file
        processed_files = []
        for file_path in excel_files:
            try:
                print(f"\n--- Processing: {os.path.relpath(file_path)} ---")
                result_file = extract_gfb_data(file_path)
                processed_files.append(result_file)
                
            except Exception as e:
                print(f"Error processing {file_path}: {str(e)}")
                continue
        
        # Summary
        print(f"\nProcessing completed!")
        print(f"Successfully processed: {len(processed_files)} out of {len(excel_files)} files")
        
        if processed_files:
            print("\nOutput files created:")
            for output_file in processed_files:
                print(f"  - {output_file}")
        
    except Exception as e:
        print(f"\nScript failed: {str(e)}")
        print("Please check that:")
        print("1. Excel files exist in current directory or subfolders")
        print("2. Files contain 'rpgBorrowing' and 'rpgRedemptions' sheets")
        print("3. Sheets have data starting from row 14-15")
        print("4. You have write permissions in the current directory")