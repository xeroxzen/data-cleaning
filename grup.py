
import pandas as pd
import numpy as np
import re
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


def custom_tokenizer(text):
    """Custom tokenizer function that handles brand names with dots and slashes."""
    return re.split(r'[\\/.]|(?<!\d)[^\w\s](?!\d)', str(text).lower())


# Read in the CSV file
df = pd.read_csv('rum.csv')

# Preprocess the Brand column
vectorizer = TfidfVectorizer(tokenizer=custom_tokenizer)
brand_vectors = vectorizer.fit_transform(df['Brand'].fillna(''))

# Group by the Brand column and create a dictionary of DataFrames
group_dict = {}
for idx, row in df.iterrows():
    if not pd.isna(row['Brand']):
        for i in range(3, 10):
            b = str(row['Brand'])
            brand_words = b.split()[:i]  # Consider the first i words of the Brand name
            brand_prefix = ' '.join(brand_words)
            for group_brand in group_dict.keys():
                g = str(group_brand)
                group_brand_words = g.split()[:i]
                group_brand_prefix = ' '.join(group_brand_words)
                if brand_prefix == group_brand_prefix:
                    group_data = group_dict[group_brand]
                    similarity = cosine_similarity(vectorizer.transform([row['Brand']]), vectorizer.transform([group_data['Brand'].iloc[0]]))[0][0]
                    if similarity >= 0.98:
                        group_dict[group_brand] = pd.concat([group_data, row.to_frame().transpose()], ignore_index=True)
                        break
            else:
                group_dict[row['Brand']] = row.to_frame().transpose()

# Create a new Excel file and add each group of duplicates to a new sheet
# with pd.ExcelWriter('new_rum.xlsx') as writer:
#     print(".....running")

#     unique_data = df.drop_duplicates(subset='Brand', keep=False)
#     unique_data.to_excel(writer, sheet_name='Unique', index=False)
#     for group, data in group_dict.items():
#         if len(data) > 1:
#             sheet_name = re.sub('[\\[\\]\\\\/:\\*\\?\\"]', '_', str(group))[:30]  # Replace invalid characters with underscores and truncate the sheet name if it's too long
#             data.to_excel(writer, sheet_name=sheet_name, index=False)
#     print("done")

with pd.ExcelWriter('new_rum.xlsx') as writer:
    print(".....running")
    unique_data = df.drop_duplicates(subset='Brand', keep=False)
    unique_data.to_excel(writer, sheet_name='Unique', index=False)
    for group, data in group_dict.items():
        if len(data) > 1:
            # sheet_name = re.sub('[\\/:\"]', '_', str(group))[:30]  # Replace invalid characters with underscores and truncate the sheet name if it's too long
            sheet_name = f'{group[:30]}'  # Truncate the sheet name if it's too long
            data.to_excel(writer, sheet_name=sheet_name, index=False)
print("done")
