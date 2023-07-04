"""
@author: Google Jr
@project: Smart Parser
"""

import argparse
import os
import pandas as pd
import re
from fuzzywuzzy import fuzz, process
import difflib


def excel_to_csv(excel_file):
    try:
        df = pd.read_excel(excel_file)

        csv_file = os.path.splitext(excel_file)[0] + '_converted.csv'
        df.to_csv(csv_file, index=False)
        print("Converted Excel file to CSV file")
    except Exception as err:
        print(f"Error occurred while converting Excel file to CSV: {str(err)}")


def clean_csv(csv_file):
    try:
        df = pd.read_csv(csv_file)
    except FileNotFoundError:
        print(f"Error: {csv_file} not found")
        return

    df = df.dropna(axis=1, how='all')

    pattern = r'(?i)\b((\d+°?\s*EAST)?\s*Hyōgo\s*(?:Dry\s*)?(?:Gin|Black Edition Gin)?\s*(?:\d+x\s*\d+\s*Tonic Water)?)\b|\b((\d+)\s*James\s*E\.\s*Pepper\s*Straight\s*(?:Whiskey|Rye|Bourbon)\s*\d*\s*(Proof)?)\b'

    df['Pattern'] = df['Brand'].apply(lambda x: re.findall(
        pattern, x)[0] if re.findall(pattern, x) else x)

    df = df.drop_duplicates(subset='Pattern')

    df = df.drop(columns=['Pattern'])

    cleaned_file = os.path.splitext(csv_file)[0] + '_cleaned.csv'
    df.to_csv(cleaned_file, index=False)
    print("Cleaned CSV file")


def csv_to_excel(csv_file):
    try:
        df = pd.read_csv(csv_file)

        brand_set = set(df['Brand'].tolist())
        fuzzy_matches = []
        for brand in brand_set:
            matches = process.extract(
                brand, brand_set, scorer=fuzz.partial_ratio, limit=None)
            for match in matches:
                if match[1] >= 90 and match[0] != brand:
                    fuzzy_matches.append(match[0])
        df['is_fuzzy_match'] = df['Brand'].isin(fuzzy_matches)
        df = df.style.apply(lambda x: ['background-color: yellow' if x['is_fuzzy_match']
                            else '' for i in x], axis=1).hide_columns(['is_fuzzy_match'])

        excel_file = os.path.splitext(csv_file)[0] + '_converted_back_to.xlsx'
        df.to_excel(excel_file, index=False)
        print("Converted CSV file to Excel file")
    except Exception as err:
        print(f"Error occurred while converting CSV file to Excel: {str(err)}")


def remove_duplicates(excel_file):
    try:
        df = pd.read_excel(excel_file)

        threshold = 0.7  # 70% match threshold... change this to a threshold of your choice

        def is_duplicate(brand, unique_brands):
            for unique_brand in unique_brands:
                similarity = difflib.SequenceMatcher(
                    None, brand, unique_brand).ratio()
                if similarity >= threshold:
                    return True
            return False

        unique_brands = []
        duplicate_indices = []

        for i, row in df.iterrows():
            brand = row['Brand']
            if is_duplicate(brand, unique_brands):
                duplicate_indices.append(i)
            else:
                unique_brands.append(brand)

        df_cleaned = df.drop(duplicate_indices)

        cleaned_file = os.path.splitext(excel_file)[0] + '_cleaned.xlsx'
        df_cleaned.to_excel(cleaned_file, index=False)
        print("Removed duplicates from Excel file")

    except Exception as err:
        print(
            f"Error occurred while removing duplicates from Excel file: {str(err)}")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Excel/CSV cleaning tool')

    parser.add_argument('--xl', metavar='EXCEL_FILE',
                        type=str, help='convert Excel file to CSV')
    parser.add_argument('--clean', metavar='CSV_FILE',
                        type=str, help='clean a CSV file or directory')
    parser.add_argument('--csv', metavar='CSV_FILE',
                        type=str, help='convert CSV file to Excel')
    parser.add_argument('--rd', metavar='EXCEL_FILE',
                        type=str, help='remove duplicates from Excel file')

    args = parser.parse_args()

    try:
        if args.xl:
            excel_to_csv(args.xl)
        elif args.clean:
            if os.path.isdir(args.clean):
                for file in os.listdir(args.clean):
                    if file.endswith('.csv'):
                        clean_csv(os.path.join(args.clean, file))
            else:
                clean_csv(args.clean)
        elif args.csv:
            csv_to_excel(args.csv)
        elif args.rd:
            remove_duplicates(args.rd)
        else:
            parser.print_help()
    except Exception as e:
        print(f"An error occurred: {str(e)}")
