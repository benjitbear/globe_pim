import pandas as pd
import re
import os

def process_excel_year_data():
    """
    Loads an Excel file, parses year data, concatenates specified columns (skipping empty cells),
    flags concatenated data exceeding 60 characters, and saves to a new CSV file with logging.
    """

    try:
        input_file = input("Enter the path to the input Excel file: ")

        if not os.path.exists(input_file):
            raise FileNotFoundError(f"File not found: {input_file}")

        if not input_file.lower().endswith(".xlsx"):
            raise ValueError("Input file must be an Excel file (.xlsx)")

        print(f"Logging: Reading Excel file: {input_file}")

        # Read Excel file
        df = pd.read_excel(input_file)

        # Identify required columns
        required_columns = ["X_veh_manufacturer", "X_manufacturer_model", "X_body_type", "X_year", "Description"]

        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Column '{col}' not found in the Excel file.")

        print("Logging: Required columns found.")

        # Strip whitespace from the year column
        df["X_year"] = df["X_year"].astype(str).str.strip()

        print("Logging: Stripped whitespace from 'X_year' column.")

        # Apply regex to parse year data
        pattern = r"\s*(\d{1,2})[/\\](\d{2,4})[-\s]+"

        def parse_year(year_str):
            match = re.search(pattern, str(year_str))
            if match:
                month = match.group(1)
                year = match.group(2)
                return f"{month}/{year}"
            else:
                return year_str  # Return original if no match

        df["Parsed_Year"] = df["X_year"].apply(parse_year)

        print("Logging: Parsed year data using regex.")

        # Concatenate specified columns with spaces, skipping empty cells
        def concatenate_with_skipping(row):
            parts = []
            for col in required_columns:
                val = row[col]
                if pd.notna(val):  # Check if cell is not NaN or None
                    parts.append(str(val))
            return ' '.join(parts)

        df["Concatenated_Data"] = df.apply(concatenate_with_skipping, axis=1)

        print("Logging: Concatenated data, skipping empty cells.")

        # Flag concatenated data exceeding 60 characters
        df["Concatenated_Data_Length"] = df["Concatenated_Data"].str.len()
        df["Exceeds_60_Characters"] = df["Concatenated_Data_Length"] > 60

        print("Logging: Flagged concatenated data exceeding 60 characters.")

        output_file = "output_description_concat.csv"

        # Handle filename collisions
        counter = 1
        while os.path.exists(output_file):
            name, ext = os.path.splitext(output_file)
            output_file = f"{name}_{counter}{ext}"
            counter += 1

        df.to_csv(output_file, index=False)

        print(f"Logging: Processed data saved to {output_file}")

    except FileNotFoundError as e:
        print(f"Error: {e}")
    except ValueError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    process_excel_year_data()