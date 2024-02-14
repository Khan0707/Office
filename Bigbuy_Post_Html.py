# -*- coding: utf-8 -*-
"""
Created on Tue Nov 21 08:24:44 2023

@author: aaqib
"""
# There are a few changes that were fixed.

# 1) The Ladies' in the title and handle was changed to Women's

# 2) The tag Category_Fashion | Accessories was changed to Category_Fashion & Accessories

# 3) The Refurbished products were deleted as we don't sell refurbished products.

import pandas as pd
import regex as re
from datetime import datetime

python_script_cleaned = r"D:\BackendData\Bigbuy\05_12_Bigbuy_CleanedDesc_GKZ_HTML.xlsx"
initial_raw_file= r"D:\BackendData\Bigbuy\05_12_Bigbuy_RawExport_GKZ.xlsx"

country_code = initial_raw_file.split('_')[-1][:2].upper()


# Read the Excel file into a DataFrame
df = pd.read_excel(python_script_cleaned)
df2= pd.read_excel(initial_raw_file)





# Define the conversion function
def convert_dimensions(input_string):
    # Match decimal numbers separated by "x" and replace commas with dots
    pattern = re.compile(r'\b(\d+(?:,\d+)?(?:\s*x\s*\d+(?:,\d+)?)*)\b')
    result = re.sub(pattern, lambda match: match.group(0).replace(',', '.'), input_string)

    # Replace patterns like "10.5 cm" with "10 cm"
    pattern_cm = re.compile(r'(\d+\.\d+) cm')
    result = re.sub(pattern_cm, lambda x: f'{float(x.group(1)):.0f} cm', result)

    pattern_decimal = re.compile(r'(\d+\.\d+)')
    result = re.sub(pattern_decimal, lambda x: f'{float(x.group(1)):.0f}', result)

    return result


# Apply the conversion function to the 'Dimensions' column
df['Body HTML'] = df['Body HTML'].apply(convert_dimensions)



# Find indices where 'Refurbished' is present in the 'Product_Description' column
refurbished_indices = df[df['Body HTML'].str.contains('Refurbished', case=False)].index

#Ladies'
# Print the indices
print("Indices with 'Refurbished':", refurbished_indices)

# Extract initial word
df['Initial_Word'] = df['Body HTML'].str.split().str[0]
# Count occurrences
word_counts = df['Initial_Word'].value_counts()
print(f"Starting Words are : {word_counts}")


def replace_specifications(text):
    # Find the first occurrence of ' <br> •' irrespective of spaces
    index = text.find('<br> •')
    
    # If the pattern is found, replace the text before it
    if index != -1:
        replaced_text = "<strong>Specifications:</strong><br>" + text[index + len('<br>•'):]#.strip()
        return replaced_text
    else:
        # If the pattern is not found, return the original text
        text #= "<strong>Highlights:</strong><br>" + text.replace('<br>',' ')#.strip()
        return text

def before_specifications(text):
    # Find the first occurrence of ' <br> •' irrespective of spaces
    index = text.find('<br> •')
    
    # If the pattern is found, replace the text before it
    if index != -1:
        replaced_text = text[:index + len('<br>•')]
        return replaced_text
    else:
        # If the pattern is not found, return the original text
        #replaced_text = "<strong>Highlights:</strong><br>" + text
        return text 
    
df['Body HTML_OLD'] = df['Body HTML'].apply(before_specifications)

df['Body HTML'] = df['Body HTML'].apply(replace_specifications)

# Extract initial word
df['Initial_Word'] = df['Body HTML'].str.split().str[0]
# Count occurrences
word_counts = df['Initial_Word'].value_counts()
print(f"Starting Words are : {word_counts}")


#df['Body HTML'].str.startswith("<strong>Specifications:</strong><br>").sum()


# Merge the DataFrames on 'Variant SKU' and replace 'Body HTML' in df2 with 'Body HTML' in df
df3 = pd.merge(df2, df[['Variant SKU', 'Body HTML']], on='Variant SKU', how='left', suffixes=('_replace', '_original'))
# Use 'Body HTML_original' as the final 'Body HTML' column in df3
df3['Body HTML'] = df3['Body HTML_original'].combine_first(df3['Body HTML_replace'])
# Drop the unnecessary columns
df3 = df3.drop(['Body HTML_replace', 'Body HTML_original'], axis=1)
# Set the column order to match df2
column_order = df2.columns
df3 = df3[column_order]

def process_df3(df):
    # Iterate over rows in df3
    for index, row in df.iterrows():
        # Extract the brand from Tags column
        brand_from_tags = row['Tags'].split(',')[0].strip().lower().replace('brand_', '').split(' ')[0]

        # Extract the brand from Title column
        brand_from_title = row['Title'].split()
        lowercase_list = [element.lower() for element in brand_from_title]

        # Check if the brand from Tags matches the brand from Title
        if brand_from_tags.lower() in lowercase_list:
            # Add "by" before the brand name in the Title column
            new_title = re.sub(r'\b{}\b'.format(re.escape(brand_from_tags)), f'by {brand_from_tags}', row['Title'], flags=re.IGNORECASE)
            df.at[index, 'Title'] = new_title.strip().title()

    # Replicate the value of Title to Handle column in lowercase using '-'
    df['Handle'] = df['Title'].str.lower().str.replace(' ', '-')
    df['Title'] = df['Title'].apply(lambda x: x.title())
    return df

df3_result = process_df3(df3)



# test=df3_result.loc[df3_result['Variant SKU']=='PIMW-94028']

# Check if 'not_update_CA' is present in Tags with the value in Vendor column

current_date = datetime.now().strftime('%d_%m')
df3_result['Tags']= df3_result['Tags'].apply(lambda x: x.replace('not_update_CA', f'{df3_result["Vendor"].iloc[0]}_{current_date}') if 'not_update_CA' in x else x)

df3_result['Tags'] =df3_result['Tags'].str.replace(r'color', 'Colour', case=False)

# Apply the modifications directly to the DataFrame
df3_result['Tags'] =df3_result['Tags'].apply(lambda tags: re.sub(r'Shipping_\d+\.\d+,', '', tags, flags=re.IGNORECASE))
# Corrected regular expression to remove patterns like "Shipping_*******" from the 'Tags' column
df3_result['Tags'] = df3_result['Tags'].apply(lambda tags: re.sub(r'Shipping_\d+,', '', tags, flags=re.IGNORECASE))


# Use regular expression to find the Type value in "Tags" and replace in the "Type" column
df3_result['Type'] = df3_result['Tags'].apply(lambda tags: re.search(r'Type_(.+)', tags, flags=re.IGNORECASE).group(1) if re.search(r'Type_(.+)', tags, flags=re.IGNORECASE) else None)

df3_result['Tags'] =df3_result['Tags'].str.replace(r'|', '&', case=False)

# Replace values in the Vendor column and store in a dictionary
vendor_mapping = {'idropship': 'GODIAU', 'vidaxl':'GOAUAD', 'wefullfill':'Vibe Geeks', 'bigbuy':'PDBB'}

if country_code in ['NZ', 'AU']:
    df3_result['Vendor'] = df3_result['Vendor'].replace(vendor_mapping)
else: 
    #tags_vendor = df3_result['Tags'].str.split(',').str[0].str.strip().str.lower().str.replace('brand_','').str.split(' ').str[0].str.upper()
    tags_vendor = df3_result['Tags'].str.extract(r'Brand_(.*?),', expand=False)
    df3_result['Vendor'] = tags_vendor.str.upper() if not tags_vendor.empty else 'GoKinzo'

#df3_result['Tags'].str.startswith('Brand_').sum()

# Step 3: Replace values in other columns
df3_result['Status'] = df3_result['Status'].replace({'Draft': 'Active'})
df3_result['Published'] = df3_result['Published'].replace({False: True})
df3_result['Published Scope'] = df3_result['Published Scope'].replace({'web': 'global'})

# Convert a set of columns from numeric to text format
columns_to_convert = ['ID', 'Variant Inventory Item ID', 'Variant ID','Variant Barcode']
df3_result[columns_to_convert] =df3_result[columns_to_convert].astype(str)





# Remove terms that come after ':'
df3_result['Title'] = df3_result['Title'].str.split(':').str.get(0)

# Cleaning 'Title' column
df3_result['Title'] = df3_result['Title'].str.title().str.strip()
df3_result['Title'] = df3_result['Title'].apply(convert_dimensions)

df3_result['Title']  = df3_result['Title'].str.replace(r'[^A-Za-z0-9\s]', '', regex=True)

df3_result['Title'] = df3_result['Title'].replace({"'S": "'s", "S'": "'s"}, regex=True)

df3_result['Title'] = df3_result['Title'].replace({"Cm": "cm", "Pcs": "pcs"}, regex=True)

df3_result['Title'] = df3_result['Title'].replace({"MenS": "Men's"}, regex=True)




def replace_digit_with_words(match):
    digit = match.group(1)
    digit_words = {
        '1': 'One',
        '2': 'Two',
        '3': 'Three',
        '4': 'Four',
        '5' : 'Five'# You can extend this dictionary for other digits
    }
    return digit_words.get(digit, digit)

# Apply the replacement to the 'Title' column
df3_result['Title'] = df3_result['Title'].apply(lambda x: re.sub(r'Playstation (\d) ', lambda match: 'Playstation ' + replace_digit_with_words(match) + ' ', x))





## Adding Package Includes for those were it is missing:

indexes_without_package_includes = df3_result[~df3_result['Body HTML'].str.contains('Package Includes:', case=False)].index
print(f"Count_indexes_without_package_includes: {len(indexes_without_package_includes)}")
print(f"indexes_without_package_includes: {indexes_without_package_includes}")

# # Split at the first digit
# split_result = re.split(r'(\d)', input_string, maxsplit=1)

# # Extract the desired part before the first digit
# result = split_result[0].strip()

def create_package(row):
    package_text = row['Title']
    if '<strong>package includes:</strong>' not in row['Body HTML'].lower() and 'units' in row['Title'].lower():
        
        match = re.search(r'(\d+)\s*Units', package_text) # Extract Quantity
        
        if match:
            extracted_number = str(int(match.group(1)))
        else:
            extracted_number= '000'
        
        text_title = re.split(r'(\d)', package_text, maxsplit=1)[0].strip() # Extract Text Content

        package_text = text_title
        highest_digit = extracted_number

        # Format the 'Package' column
        package_column = f'<br><br><strong>Package Includes:</strong><br> • {highest_digit} x {package_text}'

        return row['Body HTML'] + package_column
    
    elif '<br><br><strong>Package Includes:</strong>' not in row['Body HTML']:
        
        text_title = re.split(r'(\d)', package_text, maxsplit=1)[0].strip()
        
        package_includes_str = '<br><strong>Package Includes:</strong><br> • 1 x '
        
        row['Body HTML'] += package_includes_str + text_title
        
        return row['Body HTML']
    
df3_result['Body HTML'] = df3_result.apply(create_package, axis=1)   

# Extract initial word
df3_result['Initial_Word'] = df3_result['Body HTML'].str.split().str[0]
# Count occurrences
word_counts = df3_result['Initial_Word'].value_counts()
print(f"Starting Words are : {word_counts}")


indexes_without_package_includes = df3_result[~df3_result['Body HTML'].str.contains('Package Includes:', case=False)].index
print(f"Count_indexes_without_package_includes: {len(indexes_without_package_includes)}")
print(f"indexes_without_package_includes: {indexes_without_package_includes}")

def replace_text(html):
    if 'Highlights:' in html:
        return html.replace('<br><strong>Package', '<br><br><strong>Package')
    else:
        return html


df3_result['Body HTML'] = df3_result['Body HTML'].apply(replace_text)

#df3_result['Body HTML'].iloc[0] = df3_result['Body HTML'].iloc[0].replace("Necklace By Armani 9", "Necklace By Armani")

#Manual Replacements:
# df3_result['Body HTML'].iloc[1072] = df3_result['Body HTML'].iloc[244].replace("<br> Children deserve the best, that's why we present to you<br> Purse RCD Espanyol Blue White<br> , ideal for those who seek quality products for their little ones! Get<br> RCD Espanyol<br> ",'')

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace("REFURBISHED:",'<br>REFURBISHED:')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace("<br> • Includes: Includes the brand's case",'')
# df3_result['Body HTML'] = df3_result['Body HTML'].str.replace("<br>• Includes: Includes the brand's case",'')

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(r"`S", "'s", regex= True)

# Extract initial word
df3_result['Initial_Word'] = df3_result['Body HTML'].str.split().str[0]
# Count occurrences
word_counts = df3_result['Initial_Word'].value_counts()
print(f"Starting Words are : {word_counts}")
 
def clean_package_includes(text):
    """ Remove non-informative sequences after 'Package Includes:' """
    if '<br><strong>Package Includes:</strong><br>' in text:
        # Split the text at 'Package Includes:'
        parts = text.split('<br><strong>Package Includes:</strong><br>')
        before = parts[0]
        after = parts[1]
        # Remove patterns like 'L X X Units', 'X', 'L X X 3 Units', '6 Units' at the end of the text
        cleaned_text = re.sub(r'\bL\s*X\s*X\s*Units\b', '', after)
        cleaned_text = re.sub(r'\bX\b', '', cleaned_text)
        cleaned_text = re.sub(r'\bL\s*X\s*X\s*\d+\s*Units\b', '', cleaned_text)
        cleaned_text = re.sub(r'\b\d+\s*Units\b', '', cleaned_text)
    
        # Remove extra spaces and trailing spaces caused by the removal
        cleaned_text = re.sub(r'\s{2,}', ' ', cleaned_text).strip()

        # # Remove non-informative sequences like 'X X cm' or 'X 3 X cm'
        # cleaned_after = re.sub(r'\bX\s+[X0-9]+\s+X\b', '', after)
        
        return before + '<br><strong>Package Includes:</strong><br>' + cleaned_text
    else:
        return text


df3_result['Body HTML'] = df3_result['Body HTML'].apply(clean_package_includes)


# df3_result['Body HTML'].iloc[3] = df3_result['Body HTML'].iloc[3].replace("All products have been checked, they include the official brand guarantee and are in perfect working order.",'')

# Extract initial word
df3_result['Initial_Word'] = df3_result['Body HTML'].str.split().str[0]
# Count occurrences
word_counts = df3_result['Initial_Word'].value_counts()
print(f"Starting Words are : {word_counts}")


# Conditionally replace 'Initial_Word'
mask = df3_result['Initial_Word'].ne('<strong>Specifications:</strong><br>•')

# Apply the mask to filter rows and get value counts of 'Type'
type_counts = df3_result.loc[mask, 'Type'].value_counts()




# ### Missing Features or Highlights For Video GAmes ONly GKZ :

# Custom function to apply the transformation only where the mask is true
def transform_cell(row):
    match = re.search(r'buy\s*<br>(.*?)<br>\s*at', row['Body HTML'], re.IGNORECASE)
    captured_text = match.group(1).title() if match else ''
    
    package_includes_match = re.search(r'<br><strong>Package Includes:</strong><br>(.*?)$', row['Body HTML'], re.IGNORECASE | re.DOTALL)
    package_includes_text = package_includes_match.group(1).title() if package_includes_match else ''


    return (
        '<strong>Specifications:</strong>'
        '<br> • Brand : {0}'
        '<br> • Type : {1}'
        '<br> • Product : {2}'
        '<br><br><strong>Package Includes:</strong><br>{3}'
    ).format(row['Vendor'], row['Type'], captured_text, package_includes_text)



# Apply the custom function only where the mask is true
df3_result.loc[mask, 'Body HTML'] = df3_result[mask].apply(transform_cell, axis=1)

#df3_result['Body HTML'] = df3_result.apply(create_package, axis=1) 


#Add Bullets to all three sections:
def replace_br_between_keywords(text):
    if pd.notna(text):
        # highlight_index = text.find("<strong>Highlights:</strong>")
        # if highlight_index != -1:
        # # Find the starting index of "<strong>Features: </strong>"
        #     features_index = text.find("<strong>Features:</strong>")
        #     if features_index != -1:
        #         # Find the starting index of "<strong>Specifications:</strong>"
                specs_index = text.find("<strong>Specifications:</strong>")
                if specs_index != -1:
                    # Find the starting index of "<strong>Package Includes:</strong>"
                    package_index = text.find("<strong>Package Includes:</strong>")
                    if package_index != -1:
                        # Remove all occurrences of • between the specified keywords
                        #highlights_specs= text[highlight_index:features_index]
                        #features_to_specs = text[features_index:specs_index].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>')
                        specs_to_package = text[specs_index:package_index].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>').replace('<br> ', '<br>')
                        package_to_end = text[package_index:].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>')
                        # Replace <br> and <br> • with <br> • between the specified keywords
                        #text = text[features_index:specs_index].replace('<br> ', '<br>  •').replace('<br> •', '<br>  •') + text[specs_index:package_index].replace('<br>', '<br>  •').replace('<br> •', '<br>  •') + text[package_index:].replace('<br>', '<br>  •').replace('<br> •', '<br>  •')
                        text = (#highlights_specs +#.replace('.', '.<br>') +
                            #features_to_specs.replace('<br>', '<br>•') +
                            specs_to_package.replace('<br>', '<br>• ') +
                            package_to_end.replace('<br>', '<br>• ')
                        )

    return text


df3_result['Body HTML'] = df3_result['Body HTML'].apply(replace_br_between_keywords)
df3_result['Body HTML']=df3_result['Body HTML'].str.replace('<br>• Languages:<br>• ','<br>• Languages: ')
df3_result['Body HTML']=df3_result['Body HTML'].str.replace('<br>• Spanish<br>•','<br>• Languages: Spanish<br>•')
df3_result['Body HTML']=df3_result['Body HTML'].str.replace('<br>• Spanish English','<br>• Languages: English')
df3_result['Body HTML']=df3_result['Body HTML'].str.replace('<br>• English Spanish','<br>• Languages: English')
#df3_result['Body HTML']=df3_result['Body HTML'].str.replace('<br>• Spanish ','<br>• Languages: Spanish')

df3_result['Body HTML']= df3_result['Body HTML'].str.replace("Product :  Video Game For",'Product : ')

# remove data less bullets:
def remove_br_before_strong_and_between_br(text):
    if pd.notna(text) and isinstance(text, str):
        # Replace <br> • immediately before <strong>
        patterns_before_strong = ['<br> •<strong>', '<br>• <strong>', '<br>•<strong>']
        pattern_regex_before_strong = '|'.join(map(re.escape, patterns_before_strong))
        
        matches_before_strong = re.findall(pattern_regex_before_strong, text)
        if matches_before_strong:
            last_occurrence_before_strong = matches_before_strong[-1]
            text = text.replace(last_occurrence_before_strong, '<br><br><strong>')
        
        while re.search(pattern_regex_before_strong, text):
            text = re.sub(pattern_regex_before_strong, '<strong>', text)
        
        # Replace <br>•<br> with <br> until all occurrences are replaced
        patterns_between_br = ['<br>•<br>', '<br>• <br>']
        for pattern in patterns_between_br:
            while pattern in text:
                text = text.replace(pattern, '<br>')
        
        # Remove all occurrences of <br>• at the end
        text = re.sub(r'<br>•\s*$', '', text)
    
    return text

df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_br_before_strong_and_between_br)


# Create a mask for the condition
# condition_mask = (
#     (df3_result['Body HTML'].str.contains('<strong>Specifications:</strong><br>•')) &
#     (df3_result['Body HTML'].str.contains('<strong>Package Includes:</strong><br>')) &
#     (df3_result['Body HTML'].str.count('<br>•') <= 3)
# )

condition_mask = (
    (df3_result['Body HTML'].str.contains('<strong>Specifications:</strong><br>•')) &
    (df3_result['Body HTML'].str.contains('<strong>Package Includes:</strong><br>')) &
    (df3_result.apply(lambda row: row['Body HTML'].count('<br>•', 
        row['Body HTML'].find('<strong>Specifications:</strong><br>•'), 
        row['Body HTML'].find('<strong>Package Includes:</strong><br>')
    ), axis=1) <= 2)
)


# Define the transform_cell function
def transform_cell_v2(row):
    # Construct the new text
    new_text = (
        '<br> • Brand : {0}'
        '<br> • Type : {1}'
    ).format(row['Vendor'], row['Type'])
    
    # Find the existing text between the keywords
    existing_text = re.search(r'<strong>Specifications:</strong><br>•(.*?)<strong>Package Includes:</strong><br>', row['Body HTML'], re.DOTALL)
    
    package_includes_match = re.search(r'<br><strong>Package Includes:</strong><br>(.*?)$', row['Body HTML'], re.IGNORECASE | re.DOTALL)
    package_includes_text = package_includes_match.group(1)if package_includes_match else ''
    package_includes_text = '<br><br><strong>Package Includes:</strong><br>' + package_includes_text

    # If existing text is found, append the new text to it
    if existing_text:
        existing_text = existing_text.group(1)
        new_text = f'<strong>Specifications:</strong><br>•{existing_text}{new_text}{package_includes_text}'
    
    return new_text

# Apply the custom function only where the condition mask is true
df3_result.loc[condition_mask, 'Body HTML'] = df3_result[condition_mask].apply(transform_cell_v2, axis=1)

# df3_result['Body HTML'] = df3_result.apply(create_package, axis=1) 
df3_result['Body HTML']=df3_result['Body HTML'].str.replace('<br><br><br> •','<br> •')





def replace_text_between_keywords(df, keyword1, keyword2):
    df['Body HTML'] = df['Body HTML'].replace(f'{keyword1}.*?{keyword2}', f'{keyword2}', regex=True)

replace_text_between_keywords(df3_result, '<br>• At Hearts & Homies', '<br><br><strong>Package Includes:</strong>')
replace_text_between_keywords(df3_result, '<br>• Bring your', '<br><br><strong>Package Includes:</strong>')
replace_text_between_keywords(df3_result, '<br>• Hearts & Homies','<br><br><strong>Package Includes:</strong>')
replace_text_between_keywords(df3_result, '<br>• Keep','<br><br><strong>Package Includes:</strong>')

# Capitalise lines of all 3 section:
def capitalize_sentences(text, custom_stop_words=None):
    if custom_stop_words is None:
        custom_stop_words = set(['and', '-and', 'the', 'an', 'of', 'is', 'in', 'to', 'for', 'with', 'X', 'With','on', 'from', 'with', 'a', 'as', 'kg', 'cm', 'x', 'are', 'so', 'that'])

    start_pattern = r'<strong>Specifications:</strong>'
    end_pattern = r'<br><strong>Package Includes:</strong>'

    # Find the start and end positions
    start_pos = re.search(start_pattern, text)
    end_pos = re.search(end_pattern, text)

    if start_pos and end_pos:
        start_pos = start_pos.end()
        end_pos = end_pos.start()

        # Extract the text between the specified patterns
        highlighted_text = text[start_pos:end_pos]

        # Capitalize the first letter of each sentence excluding custom stop words
        sentences = re.split(r'(?<=[.!?])\s*', highlighted_text)
        capitalized_text = '. '.join(
            sentence.title() if sentence.split()[0].strip().lower() not in custom_stop_words else sentence
            for sentence in sentences if sentence
        )

        # Remove extra full stops after punctuation signs
        capitalized_text = re.sub(r'(?<=[.!?])\s*\.', '', capitalized_text)

        # Replace the original highlighted text with the modified version
        text = text[:start_pos] + capitalized_text + text[end_pos:].title()
    # else:
    #     print(f'Pattern Mismatch found for {text}')

    return text
 #CApitalise:
df3_result['Body HTML'] = df3_result['Body HTML'].apply(capitalize_sentences)
# 1 X
df3_result['Body HTML']= df3_result['Body HTML'].str.replace("1 X","1 x")

# 598 
index_to_drop = 598
df3_result = df3_result.drop(index_to_drop, axis=0, errors='ignore')

# Find rows with "Language : Spanish" in 'Body HTML' column
spanish_rows = df3_result['Body HTML'].str.contains('Languages: Spanish')

# Replace 'Command' with 'DELETE' in corresponding rows
df3_result.loc[spanish_rows, 'Command'] = 'DELETE'

# Print Variant SKU values for rows where 'Language : Spanish' was found
print("Variant SKU values for 'Language : Spanish' rows:")
print(df3_result.loc[spanish_rows, 'Variant SKU'].tolist())

df3_result['Title'] = df3_result['Title'].replace({" X ": " x "}, regex=True)

# df3_result['Body HTML'] = df3_result['Body HTML'].replace({" X ": " x ", " Cm ": " cm "}, regex=True)
# df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(r" Cm"," cm", regex= True)
# df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(r" Mm"," mm", regex= True)

df3_result['Body HTML'] = df3_result['Body HTML'].replace({"\s*X\s*": " x ", "\s*Cm\s*": " cm ", "\s*Mm\s*": " mm "}, regex=True)

df3_result['Body HTML'] = df3_result['Body HTML'].replace({"Men'S": "Men's", " x x box" : " x Xbox"}, regex=True)


# Define a function for string replacement
def replace_strings(value):
    if isinstance(value, str):
        value = re.sub(r"Ladies'", "Women's", value, flags=re.IGNORECASE)
        value = re.sub(r"Ladies", "Women", value, flags=re.IGNORECASE)
        value = re.sub(r"Category_Fashion \| Accessories", "Category_Fashion & Accessories", value, flags=re.IGNORECASE)
    return value

# Apply the replacement function only to string columns
df3_result = df3_result.applymap(replace_strings)



def remove_spaces(text):
    # Remove spaces before ':', ',', 'mm', 'cm', 'kg', '('
    text = re.sub(r'\s+(:|,|mm|cm|kg|\()', r'\1', text, flags=re.IGNORECASE)
    
    # Remove spaces within expressions like 'w x h x d'
    #text = re.sub(r'\s*([wWhHdDxX])\s*([wWhHdDxX])\s*([wWhHdDxX])\s*', r'\1x\2x\3', text, flags=re.IGNORECASE)
    
    return text
df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_spaces)

### HTML:
df4=df3_result[['Variant SKU','Title', 'Body HTML']]



output_file_path = initial_raw_file.replace("RawExport", "FinalCleanData")

df3_result.to_excel(output_file_path, index=False)
output_file_path2 = initial_raw_file.replace("RawExport", "ColumnData")
df4.to_excel(output_file_path2, index=False)


def check_character_count(html):
    character_count = len(html)
    status = 'Limit Crossed' if character_count > 1600 else 'Limit Not Crossed'
    status_style = 'color: red; font-weight: bold; font-style: italic;' if status == 'Limit Crossed' else ''
    character_count_style = 'color: red; font-weight: bold; font-style: italic;' if character_count > 1600 else ''
    return status, status_style, character_count_style
        
finalhtml = output_file_path.replace('xlsx', 'html')
        
with open(finalhtml, 'w', encoding='utf-8') as file:
    file.write('<!DOCTYPE html>\n<html lang="en">\n<head>\n')
    file.write('<meta charset="UTF-8">\n')
    file.write('<meta name="viewport" content="width=device-width, initial-scale=1.0">\n')
    file.write('<title>HTML Report</title>\n')
    file.write('<link rel="stylesheet" href="styles.css">')  # Link to an external stylesheet if needed
    file.write('<style>\n')
    file.write('body {\n')
    file.write('    font-family: \'Calibri\', sans-serif;\n')
    file.write('    margin: 0;\n')
    file.write('    padding: 0;\n')
    file.write('    background-color: #f4f4f4;\n')
    file.write('    color: #333;\n')
    file.write('}\n')
    file.write('main {\n')
    file.write('    max-width: 800px;\n')
    file.write('    margin: 20px auto;\n')
    file.write('    padding: 20px;\n')
    file.write('    background-color: #fff;\n')
    file.write('    box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);\n')
    file.write('}\n')
    file.write('h1, h2, h3 {\n')
    file.write('    color: #333;\n')
    file.write('}\n')
    file.write('img {\n')
    file.write('    max-width: 100%;\n')
    file.write('    height: auto;\n')
    file.write('}\n')
    file.write('p {\n')
    file.write('    line-height: 1.25;\n')  # Set line spacing to 1.0
    file.write('}\n')
    file.write('.body-html {\n')
    file.write('    font-size: normal;\n')
    file.write('    text-align: justify;\n')
    file.write('}\n')
    file.write('</style>\n')
    file.write('</head>\n<body>\n')

    # Iterate through rows of the DataFrame
    for index, row in df3_result.iterrows():
        variant_sku = row['Variant SKU']
        title = row['Title']
        body_html = row['Body HTML']
        character_count_status, status_style, character_count_style = check_character_count(body_html)

        # Write the sections to the HTML file with styling
        file.write(f'<main>\n')
        file.write(f'    <h3 style="font-size: larger;">S.No: {index} | Variant SKU: {variant_sku}</h3>\n')
        file.write(f'    <h2>Title: {title}</h2>\n')
        file.write(f'    <p style="font-size: normal; {character_count_style}">Character Count: <span style="{character_count_style}">{len(body_html)}</span> | Character Count Status: <span style="{status_style}">{character_count_status}</span></p>\n')
        file.write(f'    <div class="body-html">{body_html}</div>\n')
        file.write(f'</main>\n\n')

    file.write('</body>\n</html>')



print(f"HTML file :{finalhtml} created successfully.")