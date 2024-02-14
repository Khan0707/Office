# -*- coding: utf-8 -*-
"""
Created on Tue Nov 14 07:18:30 2023

@author: aaqib
"""

import pandas as pd
import regex as re
from datetime import datetime
import nltk


python_script_cleaned = r"D:\BackendData\Dropsite\14_11_Dropsite_CleanedDesc_NZ_HTML.xlsx"

initial_raw_file= r"D:\BackendData\Dropsite\14_11_Dropsite_RawExport_NZ.xlsx"


# Read the Excel file into a DataFrame
df = pd.read_excel(python_script_cleaned)
df2= pd.read_excel(initial_raw_file)

#add Highligts to the begining of the HTML Column:
    
df['Body HTML'] = '<strong>Highlights:</strong><br>' + df['Body HTML']

## Add Fetures instead of Key Features
df['Body HTML'] = df['Body HTML'].str.replace('key features:|key features :|key features :|features:', '<br><br><strong>Specifications:</strong>', case=False, regex=True)

# Merge the DataFrames on 'Variant SKU' and replace 'Body HTML' in df2 with 'Body HTML' in df

df3 = pd.merge(df2, df[['Variant SKU', 'Body HTML']], on='Variant SKU', how='left', suffixes=('_replace', '_original'))

# Use 'Body HTML_original' as the final 'Body HTML' column in df3
df3['Body HTML'] = df3['Body HTML_original'].combine_first(df3['Body HTML_replace'])

# Drop the unnecessary columns
df3 = df3.drop(['Body HTML_replace', 'Body HTML_original'], axis=1)

# Set the column order to match df2
column_order = df2.columns
df3 = df3[column_order]

# Modifying Title, and othe colns:

def process_df3(df):
    # Iterate over rows in df3
    for index, row in df.iterrows():
        # Extract the brand from Tags column
        brand_from_tags = row['Tags'].split(',')[0].strip().lower().replace('brand_','').split(' ')[0]

        # Extract the brand from Title column
        brand_from_title = row['Title'].split()[0].lower()

        # Check if the brand from Tags matches the brand from Title
        if brand_from_tags == brand_from_title:
            # Drop the word following 'Brand_' from all cells of Title column
            # brand_length = len('Brand_')
            df.at[index, 'Title'] = row['Title'].lower().replace(f'{brand_from_title}', '', 1).strip()

        else:
            print(f"Title, Handle, and Tags are in order for SKU {row['Variant SKU']}.")

    # Replicate the value of Title to Handle column in lowercase using '-'
    df['Handle'] = df['Title'].str.lower().str.replace(' ', '-')
    df['Title'] = df['Title'].apply(lambda x: x.title())
    return df  # Return the modified DataFrame

# Call the function with df3 and capture the result
df3_result = process_df3(df3)

# Check if 'not_update_CA' is present in Tags with the value in Vendor column

current_date = datetime.now().strftime('%d_%m')
df3_result['Tags']= df3_result['Tags'].apply(lambda x: x.replace('not_update_CA', f'{df3_result["Vendor"].iloc[0]}_{current_date}') if 'not_update_CA' in x else x)

df3_result['Tags'] =df3_result['Tags'].str.replace(r'color', 'Colour', case=False)

# Replace values in the Vendor column and store in a dictionary
vendor_mapping = {'idropship': 'GODIAU', 'vidaxl':'GOAUAD'}
df3_result['Vendor'] = df3_result['Vendor'].replace(vendor_mapping)

# Step 3: Replace values in other columns
df3_result['Status'] = df3_result['Status'].replace({'Draft': 'Active'})
df3_result['Published'] = df3_result['Published'].replace({False: True})
df3_result['Published Scope'] = df3_result['Published Scope'].replace({'web': 'global'})

# Convert a set of columns from numeric to text format
columns_to_convert = ['ID', 'Variant Inventory Item ID', 'Variant ID']
df3_result[columns_to_convert] =df3_result[columns_to_convert].astype(str)

#
# Capitalise lines of Highlights section:
def capitalize_sentences(text):
    start_pattern = r'<strong>Highlights:</strong><br>'
    end_pattern = r'<br><br><strong>Specifications:</strong><br>'

    # Find the start and end positions
    start_pos = re.search(start_pattern, text)
    end_pos = re.search(end_pattern, text)

    if start_pos and end_pos:
        start_pos = start_pos.end()
        end_pos = end_pos.start()

        # Extract the text between the specified patterns
        highlighted_text = text[start_pos:end_pos]

        # Capitalize the first letter of each sentence
        sentences = re.split(r'(?<=[.!?])\s*', highlighted_text)
        capitalized_text = '. '.join(sentence.capitalize() for sentence in sentences if sentence)

        # Remove extra full stops after punctuation signs
        capitalized_text = re.sub(r'(?<=[.!?])\s*\.', '', capitalized_text)

        # Replace the original highlighted text with the modified version
        text = text[:start_pos] + capitalized_text + text[end_pos:]

    return text
df3_result['Body HTML'] = df3_result['Body HTML'].apply(capitalize_sentences)


# import pandas as pd
# from nltk.tokenize import word_tokenize
# from nltk.corpus import words
# from nltk.corpus import stopwords
# from spellchecker import SpellChecker
# import nltk

# # nltk.download('punkt')
# # nltk.download('words')
# # nltk.download('stopwords')

# def correct_text(text):
#     # Tokenize the text into words
#     words_list = word_tokenize(text)

#     # Remove stopwords
#     stop_words = set(stopwords.words('english'))
#     words_list = [word for word in words_list if word.lower() not in stop_words]

#     # Spell correction
#     spell = SpellChecker()
#     corrected_words = [spell.correction(word) for word in words_list]

#     # Reconstruct the corrected text
#     corrected_text = ' '.join(corrected_words)

#     return corrected_text


# # Apply correction to the 'Body HTML' column
# df3_result['Body HTML'] = df3_result['Body HTML'].apply(correct_text)


