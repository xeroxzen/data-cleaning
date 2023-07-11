import argparse
import os
import pandas as pd
import re
import difflib
import shutil
from tqdm import tqdm
from colorama import init, Fore, Style

# Initialize Colorama
init()

print(Fore.RED + Style.BRIGHT + """
                      _____                           
                     | __  \                          _   
                     | |__) |_ _ _ __ ___  ___ _ __ _| |_ 
                     |  ___/ _` | '__/ __|/ _ \ '__|_   _|
                     | |  | (_| | |  \__ \  __/ |    |_|  
                     |_|   \__,_|_|  |___/\___|_|         

                {Fore.RESET}by:{Fore.RED} Andile Jaden Mbele (xeroxzen) {Fore.RESET}
                            Version: {Fore.RED}0.3{Fore.RESET}
""" + Style.RESET_ALL)


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

        # Move the original file to 'originals' folder
        original_dir = os.path.join(os.getcwd(), 'originals')
        os.makedirs(original_dir, exist_ok=True)
        shutil.move(excel_file, os.path.join(
            original_dir, os.path.basename(excel_file)))

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
                if brand.strip() not in duplicate_indices:
                    duplicate_indices.append(brand.strip())
            else:
                unique_brands.append(brand.strip())

        df_cleaned = df[~df['Brand'].isin(duplicate_indices)]

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
    parser.add_argument('path', metavar='PATH', type=str,
                        nargs='?', help='path to the Excel file')

    args = parser.parse_args()

    try:
        if args.clean:
            path = args.path
            if not path:
                print("Error: No path specified.")
                exit(1)

            if os.path.isfile(path):
                with tqdm(total=1, desc="Cleaning File", bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}{postfix}]", colour='green') as pbar:
                    clean_excel(path)
                    pbar.update(1)
            elif os.path.isdir(path):
                file_list = [os.path.join(path, file) for file in os.listdir(
                    path) if file.endswith('.xlsx')]
                for file in tqdm(file_list, desc="Cleaning Files", bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}{postfix}]", colour='green'):
                    clean_excel(os.path.join.join(path, file))
        elif args.rd:
            path = args.path
            if not path:
                print("Error: No path specified.")
                exit(1)

            if os.path.isfile(path):
                remove_duplicates(path)
            elif os.path.isdir(path):
                for file in os.listdir(path):
                    if file.endswith('.xlsx'):
                        remove_duplicates(os.path.join(path, file))
        else:
            parser.print_help()
    except Exception as e:
        print(f"An error occurred: {str(e)}")
