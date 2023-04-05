"""
@author: Google Jr
@project: Smart Parser
"""

import argparse
import os
import pandas as pd
import re
from fuzzywuzzy import fuzz, process


def excel_to_csv(excel_file):
    try:
        # Load the Excel file into a Pandas dataframe
        df = pd.read_excel(excel_file)

        csv_file = os.path.splitext(excel_file)[0] + '_converted.csv'
        df.to_csv(csv_file, index=False)
        print("Converted Excel file to CSV file")
    except Exception as err:
        print(f"Error occurred while converting Excel file to CSV: {str(err)}")


def clean_csv(csv_file):
    # Load the CSV file into a Pandas dataframe
    try:
        df = pd.read_csv(csv_file)
    except FileNotFoundError:
        print(f"Error: {csv_file} not found")
        return

    # Remove empty columns
    df = df.dropna(axis=1, how='all')

    # Remove duplicates based on the 'Brand' column and keep the row with the longest description
    df = df.sort_values('Text', key=lambda col: col.str.len(), ascending=False).drop_duplicates('Brand').reset_index(drop=True)

    # Convert 'Alcohol_Content' column values to percentages
    # Unused, ideally this was to be used to convert values like '12,5%' to '12.5%'
    def convert_percent(x):
        if isinstance(x, str) and x.endswith('%'):
            return float(x[:-1]) * 100
        else:
            return x

    # df['Alcohol_Content'] = df['Alcohol_Content'].apply(lambda x: str(int(float(x.replace(',', '.'))*100))+'%' if 
    # ',' in str(x) else str(x))
    df['Alcohol_Content'] = df['Alcohol_Content'].apply(lambda x: re.findall(r'\d+\.\d+', str(x))[0] if re.findall(r'\d+\.\d+', str(x)) else str(x))
    df['Alcohol_Content'] = df['Alcohol_Content'].apply(lambda x: str(int(float(x)*100))+'%' if '.' in str(x) else str(x))

    # Fill missing fields with data from duplicates as long as the 'Brand' column is the same, as was specified in 
    # the task
    df = df.groupby('Brand').apply(lambda x: x.ffill().bfill())

    # Save the cleaned dataframe to a new CSV file
    assert isinstance(csv_file, object)
    cleaned_file = os.path.splitext(csv_file)[0] + '_cleaned.csv'
    df.to_csv(cleaned_file, index=False)
    print("Cleaned CSV file")


def csv_to_excel(csv_file):
    try:
        # Load the CSV file into a Pandas dataframe
        df = pd.read_csv(csv_file)

        # Highlight rows with fuzzy matches in the 'Brand' column
        brand_set = set(df['Brand'].tolist())
        fuzzy_matches = []
        for brand in brand_set:
            matches = process.extract(brand, brand_set, scorer=fuzz.partial_ratio, limit=None)
            for match in matches:
                if match[1] >= 90 and match[0] != brand:
                    fuzzy_matches.append(match[0])
        df['is_fuzzy_match'] = df['Brand'].isin(fuzzy_matches)
        df = df.style.apply(lambda x: ['background-color: yellow' if x['is_fuzzy_match'] else '' for i in x], axis=1).hide_columns(['is_fuzzy_match'])

        # Save the dataframe to a new Excel file
        excel_file = os.path.splitext(csv_file)[0] + '_converted_back_to.xlsx'
        df.to_excel(excel_file, index=False)
        print("Converted CSV file to Excel file")
    except Exception as err:
        print(f"Error occurred while converting CSV file to Excel: {str(err)}")



if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Excel/CSV cleaning tool')

    parser.add_argument('--xl', metavar='EXCEL_FILE', type=str, help='convert Excel file to CSV')
    parser.add_argument('--clean', metavar='CSV_FILE', type=str, help='clean a CSV file or directory')
    parser.add_argument('--csv', metavar='CSV_FILE', type=str, help='convert CSV file to Excel')

    args = parser.parse_args()

    try:
        if args.xl:
            excel_to_csv(args.xl)
        elif args.clean:
            if os.path.isdir(args.clean):
                # Clean all CSV files in a directory
                for file in os.listdir(args.clean):
                    if file.endswith('.csv'):
                        clean_csv(os.path.join(args.clean, file))
            else:
                # Clean a single CSV file
                clean_csv(args.clean)
        elif args.csv:
            csv_to_excel(args.csv)
        else:
            parser.print_help()
    except Exception as e:
        print(f"An error occurred: {str(e)}")
