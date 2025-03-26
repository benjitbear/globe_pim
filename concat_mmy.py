body_type_mapping = {
    'SEDAN': 'SDN',
    'HATCHBACK': 'HBK',
    'WAGON': 'WGN',
    'COUPE': 'CPE',
    'CABRIOLET': 'CBL',
    'ROADSTER': 'RDS',
    'SUV': 'SUV',
    'SPORTSBACK': 'SBK',
    'SPORTS BACK': 'SBK',
    'UTE': 'UTE',
    'LIFTBACK': 'LBK',
    'LIFT BACK': 'LBK',
    'VAN': 'VAN',
    'CONVERTIBLE': 'CNV',
    'CONV': 'CNV',
    'QUATTRO': 'QTR',
    'TRUCK': 'TRK',
    'BUS': 'BUS'
}

import pandas as pd
import argparse
import logging
import numpy as np  # Import numpy for checking NaN

# Set up basic logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def process_excel_data(excel_file, print_to_terminal=False, verbose=False):
    """
    Processes an Excel file to generate a new spreadsheet or print to the terminal
    with mapped body types and the number of doors. Handles empty or non-numeric
    values in the 'doors' column gracefully. Skips rows with missing core data
    and logs a warning. Handles partial matches in the 'body_type' column.
    Ensures all components of the description are strings.

    Args:
        excel_file (str): Path to the input Excel file.
        print_to_terminal (bool, optional): If True, prints the output to the
                                            terminal instead of saving to a file.
                                            Defaults to False.
        verbose (bool, optional): If True, enables verbose output and logging.
                                   Defaults to False.
    """
    if verbose:
        logging.getLogger().setLevel(logging.DEBUG)
        logging.debug("Verbose mode enabled.")

    try:
        logging.info(f"Loading Excel file: {excel_file}")
        df = pd.read_excel(excel_file)
        logging.debug(f"Excel file loaded successfully. Columns: {df.columns.tolist()}")

        output_data = []
        exo_part_number_columns = [col for col in df.columns if 'EXO PART NUMBER' in col]
        logging.debug(f"Found EXO PART NUMBER columns: {exo_part_number_columns}")

        for index, row in df.iterrows():
            if verbose:
                logging.debug(f"Processing row index: {index}")
                logging.debug(f"Row data: {row.to_dict()}")
            try:
                make = row.get('make')
                model = row.get('model')
                year_text = row.get('year_text')
                body_type_raw = str(row.get('body_type', '')).upper()
                doors = row.get('doors')
                if any(pd.isna(val) for val in [make, model, year_text, body_type_raw]):
                    logging.warning(f"Skipping row {index} due to missing required data in 'make', 'model', 'year_text', or 'body_type'.")
                    if verbose:
                        logging.debug(f"Row data: {row.to_dict()}")
                    continue # Skip the row if any of the core columns (excluding doors) are missing
            except Exception as e:
                logging.error(f"Error accessing row data at index {index}: {e}")
                continue # Still continue to the next row for other unexpected errors

            body_type_mapped = body_type_raw
            for key, value in body_type_mapping.items():
                if key.upper() in body_type_raw:
                    body_type_mapped = body_type_mapped.replace(key.upper(), value)
                    if verbose:
                        logging.debug(f"Mapped '{key}' to '{value}' in '{body_type_raw}', resulting in '{body_type_mapped}'")

            make_str = str(make) if pd.notna(make) else ''
            model_str = str(model) if pd.notna(model) else ''
            year_text_str = str(year_text) if pd.notna(year_text) else ''

            description_parts = [make_str, model_str, year_text_str]
            if pd.notna(doors):
                try:
                    doors_int = int(doors)
                    description_parts.append(f"{doors_int}DR")
                except (ValueError, TypeError):
                    logging.warning(f"Could not convert 'doors' value '{doors}' to integer for description in row {index}. Skipping doors in description.")
                    if verbose:
                        logging.debug(f"Row data: {row.to_dict()}")
            description_parts.append(body_type_mapped)

            description = ' '.join(filter(None, description_parts)).strip()

            if verbose:
                logging.debug(f"Generated description: {description}")

            for exo_col in exo_part_number_columns:
                stockcode = row[exo_col]
                if pd.notna(stockcode):
                    output_data.append({
                        'STOCKCODE': stockcode,
                        'MAKE': make if pd.notna(make) else '',
                        'MODEL': model if pd.notna(model) else '',
                        'YEAR': year_text if pd.notna(year_text) else '',
                        'DOORS': int(doors) if pd.notna(doors) and pd.notna(int(doors)) else '',
                        'BODY TYPE': body_type_mapped if body_type_mapped.strip() else '',
                        'DESCRIPTION': description
                    })
                    if verbose:
                        logging.debug(f"Extracted STOCKCODE '{stockcode}' from column '{exo_col}'")

        output_df = pd.DataFrame(output_data)
        logging.info(f"Generated output DataFrame with {len(output_df)} rows.")
        if verbose:
            logging.debug(f"Output DataFrame head:\n{output_df.head().to_string()}")

        if print_to_terminal:
            print("\n--- Output (Terminal) ---")
            print(output_df.to_string(na_rep='')) # Use na_rep to replace NaN with empty string for terminal output
        else:
            output_file = 'processed_data.xlsx'
            output_df.to_excel(output_file, index=False, na_rep='') # Use na_rep to replace NaN with empty string in Excel
            logging.info(f"Processed data saved to {output_file}")

    except FileNotFoundError:
        logging.error(f"Error: Excel file not found at the specified path: {excel_file}")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process Excel data and map body types.")
    parser.add_argument("excel_file", help="Path to the input Excel file.")
    parser.add_argument("--print", action="store_true", help="Print output to terminal instead of saving to a file.")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose output and logging.")

    args = parser.parse_args()

    process_excel_data(args.excel_file, args.print, args.verbose)