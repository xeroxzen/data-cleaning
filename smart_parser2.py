'''
@author: Google Jr
@project: Smart Parser
'''

import argparse
import os
import pandas as pd

def excel_to_csv(excel_file):
    # Load the Excel file into a Pandas dataframe
    df = pd.read_excel(excel_file)

    # Save the dataframe to a CSV file
    csv_file = os.path.splitext(excel_file)[0] + '.csv'
    df.to_csv(csv_file, index=False)

def clean_csv(csv_file):
    # Load the CSV file into a Pandas dataframe
    df = pd.read_csv(csv_file)

    # Remove empty columns
    df = df.dropna(axis=1, how='all')

    # Remove duplicates based on the 'Brand' column and keep the row with the longest description
    df = df.sort_values('Text', key=lambda col: col.str.len(), ascending=False).drop_duplicates('Brand').reset_index(drop=True)

    # Convert 'Alcohol_Content' column values to percentages
    df['Alcohol_Content'] = df['Alcohol_Content'].apply(lambda x: str(int(float(x)*100))+'%' if '.' in str(x) else str(x))

    # Fill missing fields with data from duplicates
    df = df.groupby('Brand').apply(lambda x: x.ffill().bfill())


    # Save the cleaned dataframe to a new CSV file
    cleaned_csv_file = os.path.splitext(csv_file)[0] + '_cleaned.csv'
    df.to_csv(cleaned_csv_file, index=False)

def csv_to_excel(csv_file):
    # Load the CSV file into a Pandas dataframe
    df = pd.read_csv(csv_file)

    # Save the dataframe to an Excel file
    excel_file = os.path.splitext(csv_file)[0] + '.xlsx'
    df.to_excel(excel_file, index=False)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Convert Excel to CSV, clean CSV, or convert CSV to Excel')
    parser.add_argument('--xl', metavar='EXCEL_FILE', type=str, help='Convert Excel to CSV')
    parser.add_argument('--clean', metavar='CSV_FILE', type=str, help='Clean CSV file or directory')
    parser.add_argument('--csv', metavar='CSV_FILE', type=str, help='Convert CSV to Excel')
    args = parser.parse_args()

    if args.xl:
        excel_to_csv(args.xl)
    elif args.clean:
        if os.path.isfile(args.clean):
            clean_csv(args.clean)
        elif os.path.isdir(args.clean):
            for file_name in os.listdir(args.clean):
                if file_name.endswith('.csv'):
                    file_path = os.path.join(args.clean, file_name)
                    clean_csv(file_path)
        else:
            print('Error: ' + args.clean + ' is not a valid file or directory')
    elif args.csv:
        csv_to_excel(args.csv)
    else:
        parser.print_help()
