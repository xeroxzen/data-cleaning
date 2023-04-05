import pandas as pd
import sys

def csv_to_excel(csv_file):
    # Read the CSV file into a DataFrame
    df = pd.read_csv(csv_file)

    # Write the DataFrame to an Excel file
    output_file = csv_file.replace(".csv", ".xlsx")
    df.to_excel(output_file, index=False)

    print(f"CSV file '{csv_file}' converted to Excel file '{output_file}'")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python csv_to_excel.py <csv_file>")
        sys.exit(1)

    csv_file = sys.argv[1]
    csv_to_excel(csv_file)
