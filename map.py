"""
GFB Data Extraction - Config-Based Mapper

This script extracts data from Excel files using a JSON configuration file
that maps each output column to its source location by searching for text patterns.

The OUTPUT ORDER is fixed and never changes.
The JSON config stores WHERE to find each measure in the source Excel.
"""

import pandas as pd
import os
import json
from datetime import datetime
import numpy as np


def load_config(config_file="gfb_config.json"):
    """Load the JSON configuration file"""
    if not os.path.exists(config_file):
        raise FileNotFoundError(
            f"Config file not found: {config_file}\n"
            f"Please run build_config.py first to generate the configuration."
        )

    with open(config_file, 'r', encoding='utf-8') as f:
        config = json.load(f)

    print(f"Loaded config: {config_file}")
    print(f"  Version: {config.get('version', 'unknown')}")
    print(f"  Created: {config.get('created', 'unknown')}")
    print(f"  Source: {config.get('source_file', 'unknown')}")

    return config


def format_number(value):
    """Format numbers with commas as thousand separators"""
    if pd.isna(value) or value is None:
        return None

    try:
        num_value = float(value)
        if num_value == int(num_value):
            return int(num_value)
        else:
            return num_value
    except (ValueError, TypeError):
        return value


def extract_gfb_data_with_config(source_file, config):
    """
    Extract and map data from Excel file using the JSON configuration

    Args:
        source_file: Path to source Excel file
        config: Configuration dictionary loaded from JSON

    Returns:
        Path to generated output file
    """
    print(f"\nProcessing: {source_file}")
    print("=" * 70)

    # Read sheets
    print("\n1. Reading Excel sheets...")
    borrowing_df = pd.read_excel(source_file, sheet_name='rpgBorrowing', header=None)
    redemption_df = pd.read_excel(source_file, sheet_name='rpgRedemptions', header=None)
    print(f"   Borrowing: {borrowing_df.shape[0]} rows x {borrowing_df.shape[1]} cols")
    print(f"   Redemption: {redemption_df.shape[0]} rows x {redemption_df.shape[1]} cols")

    # Get configuration
    borr_config = config['borrowing_sheet']
    redem_config = config['redemption_sheet']

    borr_date_row = borr_config['date_row']
    redem_date_row = redem_config['date_row']

    # Extract dates
    print("\n2. Extracting dates...")
    date_row_data = borrowing_df.iloc[borr_date_row, 1:].dropna()

    dates = []
    for date_val in date_row_data:
        if pd.notna(date_val):
            if isinstance(date_val, datetime):
                dates.append(date_val.strftime('%Y-%m'))
            else:
                try:
                    parsed_date = pd.to_datetime(date_val)
                    dates.append(parsed_date.strftime('%Y-%m'))
                except:
                    dates.append(str(date_val))

    print(f"   Found {len(dates)} date columns")
    print(f"   Date range: {dates[0] if dates else 'N/A'} to {dates[-1] if dates else 'N/A'}")

    # Extract borrowing data using config mappings
    print("\n3. Extracting borrowing data using config mappings...")
    borrowing_data = []

    for i, measure in enumerate(borr_config['measures'], 1):
        row_idx = measure['source_row']

        if row_idx is not None and row_idx < len(borrowing_df):
            row_data = borrowing_df.iloc[row_idx, 1:len(dates)+1].values
            borrowing_data.append(row_data)

            if i <= 5 or i > 50:  # Show first 5 and last few
                label = measure['source_label'][:50] if measure['source_label'] else "N/A"
                print(f"   Col {i:2d}: Row {row_idx+1:3d} -> {label}")
        else:
            # No mapping found - use NaN
            borrowing_data.append([np.nan] * len(dates))
            print(f"   Col {i:2d}: NOT MAPPED -> Using NaN")

    # Extract redemption data using config mappings
    print("\n4. Extracting redemption data using config mappings...")
    redemption_data = []

    for i, measure in enumerate(redem_config['measures'], 1):
        row_idx = measure['source_row']

        if row_idx is not None and row_idx < len(redemption_df):
            row_data = redemption_df.iloc[row_idx, 1:len(dates)+1].values
            redemption_data.append(row_data)

            if i <= 5 or i > 50:  # Show first 5 and last few
                label = measure['source_label'][:50] if measure['source_label'] else "N/A"
                print(f"   Col {i:2d}: Row {row_idx+1:3d} -> {label}")
        else:
            # No mapping found - use NaN
            redemption_data.append([np.nan] * len(dates))
            print(f"   Col {i:2d}: NOT MAPPED -> Using NaN")

    # Prepare output headers (fixed order from config)
    print("\n5. Preparing output structure...")

    borrowing_codes = [m['code'] for m in borr_config['measures']]
    borrowing_descriptions = [m['description'] for m in borr_config['measures']]

    redemption_codes = [m['code'] for m in redem_config['measures']]
    redemption_descriptions = [m['description'] for m in redem_config['measures']]

    all_codes = [None] + borrowing_codes + redemption_codes
    all_descriptions = [None] + borrowing_descriptions + redemption_descriptions

    # Combine data
    all_data = borrowing_data + redemption_data

    # Create output rows
    print("\n6. Building output DataFrame...")
    output_data = []

    # Header rows
    output_data.append(all_codes)
    output_data.append(all_descriptions)

    # Data rows
    for i, date in enumerate(dates):
        row = [date]  # Start with date

        # Add data for each measure
        for measure_idx in range(len(all_data)):
            if i < len(all_data[measure_idx]):
                value = all_data[measure_idx][i]
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

    # Format numbers
    print("\n7. Formatting numbers...")
    formatted_data = []
    for row in output_data:
        formatted_row = []
        for i, cell in enumerate(row):
            if i == 0 or len(formatted_data) < 2:  # Date column or header rows
                formatted_row.append(cell)
            else:
                formatted_row.append(format_number(cell))
        formatted_data.append(formatted_row)

    output_df = pd.DataFrame(formatted_data)

    # Generate output filename with output folder and timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    source_basename = os.path.splitext(os.path.basename(source_file))[0]

    # Create output folder if it doesn't exist
    output_folder = "output"
    os.makedirs(output_folder, exist_ok=True)

    output_filename = os.path.join(output_folder, f"GFB_DATA_{source_basename}_{timestamp}.xlsx")

    # Save to Excel
    print(f"\n8. Saving to {output_filename}...")

    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        output_df.to_excel(writer, sheet_name='DATA', index=False, header=False)

        # Apply number formatting
        worksheet = writer.sheets['DATA']

        for row_num in range(3, len(formatted_data) + 1):
            for col_num in range(2, len(all_codes) + 1):
                cell = worksheet.cell(row=row_num, column=col_num)
                if cell.value is not None and isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0'

    print(f"\nSUCCESS! Created {output_filename}")
    print("=" * 70)
    print(f"  Data rows: {len(formatted_data)-2}")
    print(f"  Measure columns: {len(all_codes)-1} (54 BORR + 54 REDEM)")
    print(f"  Date range: {dates[0] if dates else 'N/A'} to {dates[-1] if dates else 'N/A'}")
    print("=" * 70)

    return output_filename


def find_excel_files():
    """Scan for valid Excel files"""
    excel_files = []
    current_dir = os.getcwd()

    print("Scanning for Excel files...")

    for root, dirs, files in os.walk(current_dir):
        for file in files:
            if file.endswith('.xlsx') and not file.startswith('~') and not file.startswith('GFB_DATA_'):
                file_path = os.path.join(root, file)

                try:
                    excel_file = pd.ExcelFile(file_path)
                    sheet_names = excel_file.sheet_names

                    if 'rpgBorrowing' in sheet_names and 'rpgRedemptions' in sheet_names:
                        excel_files.append(file_path)
                        rel_path = os.path.relpath(file_path, current_dir)
                        print(f"  Found: {rel_path}")

                    excel_file.close()

                except Exception:
                    continue

    return excel_files


if __name__ == "__main__":
    print("=" * 70)
    print("GFB Data Extraction - Config-Based Mapper")
    print("=" * 70)

    try:
        # Load configuration
        config = load_config("gfb_config.json")

        # Find Excel files
        excel_files = find_excel_files()

        if not excel_files:
            print("\nERROR: No valid Excel files found!")
            print("Looking for files with 'rpgBorrowing' and 'rpgRedemptions' sheets")
            exit(1)

        print(f"\nFound {len(excel_files)} valid Excel file(s)")

        # Process each file
        processed_files = []
        for file_path in excel_files:
            try:
                result_file = extract_gfb_data_with_config(file_path, config)
                processed_files.append(result_file)

            except Exception as e:
                print(f"\nERROR processing {file_path}: {str(e)}")
                import traceback
                traceback.print_exc()
                continue

        # Summary
        print("\n" + "=" * 70)
        print("Processing Summary")
        print("=" * 70)
        print(f"Successfully processed: {len(processed_files)} out of {len(excel_files)} files")

        if processed_files:
            print("\nOutput files created:")
            for output_file in processed_files:
                print(f"  - {output_file}")

    except FileNotFoundError as e:
        print(f"\nERROR: {e}")
        print("\nPlease run build_config.py first to generate the configuration file.")
        exit(1)
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        exit(1)
