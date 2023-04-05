import argparse
import pandas as pd

def clean_csv(csv_file):
    # Load the CSV file into a Pandas dataframe
    df = pd.read_csv(csv_file)

    # Remove empty columns
    df = df.dropna(axis=1, how='all')

    # Remove duplicates based on the 'Brand' column and keep the row with the longest description
    df = df.sort_values('Text', key=lambda col: col.str.len(), ascending=False).drop_duplicates('Brand').reset_index(drop=True)

    # Fill missing fields with data from duplicates
    df = df.groupby('Brand').apply(lambda x: x.ffill().bfill())


    # Convert 'Alcohol_Content' column to proper percentages
    df['Alcohol_Content'] = df['Alcohol_Content'].apply(lambda x: str(int(float(x.rstrip('%\xa0')) * 100)) + '%' if '.' in str(x) else x)

    # Save the cleaned dataframe to a new CSV file, in the same directory with the \
    # same name appended with _cleaned before the file extension    
    df.to_csv(csv_file.split('.')[0] + '_cleaned.csv', index=False)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('csv_file', type=str, help='CSV file to clean')
    args = parser.parse_args()
    
    clean_csv(args.csv_file)
