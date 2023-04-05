import pandas as pd
import sys

def excel_to_csv(excel_file):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(excel_file)

    # Write the DataFrame to a CSV file
    output_file = excel_file.replace(".xlsx", ".csv")
    df.to_csv(output_file, index=False)

    print(f"Excel file '{excel_file}' converted to CSV file '{output_file}'")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python excel_to_csv.py <excel_file>")
        sys.exit(1)

    excel_file = sys.argv[1]
    excel_to_csv(excel_file)
