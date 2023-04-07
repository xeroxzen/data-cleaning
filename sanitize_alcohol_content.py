import pandas as pd
import os
import sys
import re

def sanitize_alcohol_content(csv_file):
    # Load the CSV file into a Pandas dataframe
    df = pd.read_csv(csv_file)

    # Define a regular expression pattern to match the alcohol content values
    pattern = r'(\d+(?:[.,]\d+)?)(?:%)?'

    # Iterate over the 'Alcohol_Content' column and sanitize its values
    for i, value in enumerate(df['Alcohol_Content']):
        # Check if the value is a string
        if isinstance(value, str):
            # Match the value against the regular expression pattern
            match = re.match(pattern, value.strip())
            if match:
                # Convert the matched value to a float and store it in the dataframe
                df.at[i, 'Alcohol_Content'] = float(match.group(1).replace(',', '.'))
            else:
                # If the value does not match the pattern, set it to NaN
                df.at[i, 'Alcohol_Content'] = float('nan')
        elif isinstance(value, (int, float)):
            # If the value is already a numeric type, leave it as is
            pass
        else:
            # If the value is neither a string nor a numeric type, set it to NaN
            df.at[i, 'Alcohol_Content'] = float('nan')

    # Save the cleaned dataframe to a new CSV file
    cleaned_csv_file = os.path.splitext(csv_file)[0] + '_cleaned.csv'
    df.to_csv(cleaned_csv_file, index=False)

    return cleaned_csv_file

if __name__ == '__main__':
    # Get the CSV file location from the command line arguments
    csv_file = sys.argv[1]

    # Sanitize the alcohol content column and save the cleaned CSV file
    cleaned_csv_file = sanitize_alcohol_content(csv_file)

    print(f"Cleaned CSV file saved to {cleaned_csv_file}")

