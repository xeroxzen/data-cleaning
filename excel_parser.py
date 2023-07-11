import argparse
import os
import pandas as pd
import re
import difflib
from tqdm import tqdm


def clean_excel(excel_file):
    try:
        df = pd.read_excel(excel_file)

        df = df.dropna(axis=1, how='all')

        df = df.sort_values('Text', key=lambda col: col.str.len(
        ), ascending=False).drop_duplicates('Brand').reset_index(drop=True)

        df['Alcohol_Content'] = df['Alcohol_Content'].apply(lambda x: re.findall(
            r'\d+\.\d+', str(x))[0] if re.findall(r'\d+\.\d+', str(x)) else str(x))
        df['Alcohol_Content'] = df['Alcohol_Content'].apply(
            lambda x: str(int(float(x)*100))+'%' if '.' in str(x) else str(x))

        df = df.groupby('Brand').apply(lambda x: x.ffill().bfill())

        cleaned_file = os.path.splitext(excel_file)[0] + '_cleaned.xlsx'
        df.to_excel(cleaned_file, index=False)
        print("Cleaned Excel file")

    except Exception as err:
        print(
            f"Error occurred while cleaning Excel file: {str(err)}")


def remove_duplicates(excel_file):
    try:
        df = pd.read_excel(excel_file)

        threshold = 0.7  # 70% match threshold... change this to a threshold of your choice

        def is_duplicate(brand, unique_brands):
            for unique_brand in unique_brands:
                similarity = difflib.SequenceMatcher(
                    None, brand.strip(), unique_brand.strip()).ratio()
                if similarity >= threshold:
                    return True
            return False

        unique_brands = []
        duplicate_indices = []

        for i, row in tqdm(df.iterrows(), total=len(df), desc="Processing", bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}{postfix}]", colour='green'):
            brand = row['Brand']
            if is_duplicate(brand, unique_brands):
                # Get the index of the existing unique row
                unique_index = unique_brands.index(brand.strip())
                # Update the existing unique row with data from duplicates
                unique_row = df.loc[unique_index]
                for col in df.columns:
                    if pd.isnull(unique_row[col]) and not pd.isnull(row[col]):
                        unique_row[col] = row[col]
            else:
                unique_brands.append(brand.strip())
                # Get the index of the newly added unique row
                unique_index = unique_brands.index(brand.strip())
                # Update the corresponding row in duplicate_indices list
                for idx, duplicate_idx in enumerate(duplicate_indices):
                    if duplicate_idx > unique_index:
                        duplicate_indices[idx] += 1

        df_cleaned = df.drop(duplicate_indices)

        cleaned_file = os.path.splitext(excel_file)[0] + '_rmdup.xlsx'
        df_cleaned.to_excel(cleaned_file, index=False)
        print(f"Removed duplicates from Excel file: {cleaned_file}")

    except Exception as err:
        print(
            f"Error occurred while removing duplicates from Excel file: {str(err)}")




if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Excel cleaning tool')

    parser.add_argument('--clean', metavar='EXCEL_FILE',
                        type=str, help='clean an Excel file')
    parser.add_argument('--rd', metavar='EXCEL_FILE',
                        type=str, help='remove duplicates from Excel file')

    args = parser.parse_args()

    try:
        if args.clean:
            clean_excel(args.clean)
        elif args.rd:
            remove_duplicates(args.rd)
        else:
            parser.print_help()
    except Exception as e:
        print(f"An error occurred: {str(e)}")
