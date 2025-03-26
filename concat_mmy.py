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

# Set up basic logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def process_excel_data(excel_file, print_to_terminal=False, verbose=False):
    """
    Processes an Excel file to generate a new spreadsheet or print to the terminal
    with mapped body types and the number of doors.
    Handles partial matches in the 'body_type' column.

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
                make = row['make']
                model = row['model']
                year_text = row['year_text']
                body_type_raw = str(row['body_type']).upper()
                doors = row['doors']
            except KeyError as e:
                logging.error(f"Error accessing column: {e}. Please ensure the Excel file has columns named 'make', 'model', 'year_text', 'body_type', and 'doors'.")
                raise  # Re-raise the exception to stop processing

            body_type_mapped = body_type_raw
            for key, value in body_type_mapping.items():
                if key.upper() in body_type_raw:
                    body_type_mapped = body_type_mapped.replace(key.upper(), value)
                    if verbose:
                        logging.debug(f"Mapped '{key}' to '{value}' in '{body_type_raw}', resulting in '{body_type_mapped}'")

            description = f"{make} {model} {year_text} {doors}DR {body_type_mapped}"
            if verbose:
                logging.debug(f"Generated description: {description}")

            for exo_col in exo_part_number_columns:
                stockcode = row[exo_col]
                if pd.notna(stockcode):
                    output_data.append({
                        'STOCKCODE': stockcode,
                        'MAKE': make,
                        'MODEL': model,
                        'YEAR': year_text,
                        'DOORS': doors,
                        'BODY TYPE': body_type_mapped,
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
            print(output_df.to_string())
        else:
            output_file = 'processed_data.xlsx'
            output_df.to_excel(output_file, index=False)
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