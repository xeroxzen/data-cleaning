import argparse
import os
import pandas as pd
import re
import difflib


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

        cleaned_file = os.path.splitext(excel_file)[0] + '_rmdup.xlsx'
        df_cleaned.to_excel(cleaned_file, index=False)
        print("Removed duplicates from Excel file")

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
