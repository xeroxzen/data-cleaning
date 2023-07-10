import argparse
import os
import pandas as pd
import re
import difflib
import shutil
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
        print(f"Cleaned Excel file: {cleaned_file}")

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

        for i, row in tqdm(df.iterrows(), total=len(df), desc="Processing", bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}{postfix}]", colour='green'):
            brand = row['Brand']
            if is_duplicate(brand, unique_brands):
                duplicate_indices.append(i)
            else:
                unique_brands.append(brand)

        df_cleaned = df.drop(duplicate_indices)

        cleaned_file = os.path.splitext(excel_file)[0] + '_rmdup.xlsx'
        df_cleaned.to_excel(cleaned_file, index=False)
        print(f"Removed duplicates from Excel file: {cleaned_file}")

    except Exception as err:
        print(
            f"Error occurred while removing duplicates from Excel file: {str(err)}")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Excel cleaning tool')

    parser.add_argument('--clean', action='store_true',
                        help='clean an Excel file or directory')
    parser.add_argument('--rd', action='store_true',
                        help='remove duplicates from Excel file or directory')

    args = parser.parse_args()

    try:
        if args.clean:
            path = input("Enter the path to the Excel file or directory: ")
            if os.path.isfile(path):
                file_list = [path]
            elif os.path.isdir(path):
                file_list = [os.path.join(path, file) for file in os.listdir(
                    path) if file.endswith('.xlsx')]
            else:
                print(f"Error: Path '{path}' not found.")
                exit(1)

            original_dir = os.path.join(os.getcwd(), 'originals')
            os.makedirs(original_dir, exist_ok=True)

            for file in tqdm(file_list, desc="Processing", bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}{postfix}]", colour='green'):
                shutil.move(file, original_dir)
                clean_excel(os.path.join(original_dir, os.path.basename(file)))
        elif args.rd:
            path = input("Enter the path to the Excel file or directory: ")
            if os.path.isfile(path):
                file_list = [path]
            elif os.path.isdir(path):
                file_list = [os.path.join(path, file) for file in os.listdir(
                    path) if file.endswith('.xlsx')]
            else:
                print(f"Error: Path '{path}' not found.")
                exit(1)

            original_dir = os.path.join(os.getcwd(), 'originals')
            os.makedirs(original_dir, exist_ok=True)

            for file in tqdm(file_list, desc="Processing", bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}{postfix}]", colour='green'):
                shutil.move(file, original_dir)
                remove_duplicates(os.path.join(
                    original_dir, os.path.basename(file)))
        else:
            parser.print_help()
    except Exception as e:
        print(f"An error occurred: {str(e)}")
