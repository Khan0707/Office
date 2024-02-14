# -*- coding: utf-8 -*-
"""
Created on Wed Nov 22 07:30:55 2023

@author: aaqib
"""

import pandas as pd
import numpy as np
import regex as re
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")#, category=DeprecationWarning)
import time

import language_tool_python

initial_raw_file="D:\BackendData\IBCRAFT\IBACRAFT_AK.xlsx"

# Start the timer
start_time = time.time()

df= pd.read_excel(initial_raw_file)
# Set the second row as column names
df.columns = df.iloc[1].str.title()

# df['Category:Name'].unique()
# len(df['Product_Type'].unique())
# count= df.groupby('Generic_Keywords')['Variant SKU'].nunique()

# Drop the first two rows
df = df.drop([0, 1])
df = df.reset_index(drop=True)
# Assuming 'Item SKU' is the current column name
df.rename(columns={'Item_Sku': 'Variant SKU', 'Feed_Product_Type':'Product_Type'}, inplace=True)

# Sample replacement dictionary
replacement_dict = {
    'wallart': 'Wall Art',
    'ruler': 'Ruler',
    'templatestencil': 'Template Stencil',
    'officeproducts': 'Office Products',
    'writinginstruments': 'Writing Instruments',
    'shoes': 'Shoes'
}
################################################################# TRACE EASE: #########################
df=df.loc[df['Brand_Name']=='Traceease']

# Assuming 'your_column' is the column you want to replace values in
df['Product_Type'] = df['Product_Type'].replace(replacement_dict)

taxonomy_dict = {
    'Ruler': [
        'Arts & Entertainment, Hobbies & Creative Arts, Arts & Crafts, Art & Crafting Tools, Craft Measuring & Marking Tools, Textile Art Gauges & Rulers'#,
        #'Hardware, Tools, Measuring Tools & Sensors, Rulers'
    ],
    'Wall Art': ['Home & Garden, Decor, Artwork, Posters, Prints & Visual Artwork'],
    'Template Stencil': [
        # 'Arts & Entertainment, Hobbies & Creative Arts, Arts & Crafts, Art & Crafting Tools, Craft Measuring & Marking Tools, Stencil Machines',
        'Arts & Entertainment, Hobbies & Creative Arts, Arts & Crafts, Art & Crafting Tools, Craft Measuring & Marking Tools, Stencils & Die Cuts'
    ],
    'Writing Instruments': ['Office Supplies, Office Instruments, Writing & Drawing Instruments, Multifunction Writing Instruments'],
    'Office Products': [
        #  'Arts & Entertainment, Hobbies & Creative Arts, Arts & Crafts, Art & Crafting Tools, Craft Measuring & Marking Tools, Stencil Machines',
        'Arts & Entertainment, Hobbies & Creative Arts, Arts & Crafts, Art & Crafting Tools, Craft Measuring & Marking Tools, Stencils & Die Cuts'
    ],
   'Shoes' :['Apparel & Accessories > Costumes & Accessories , Costume Shoes']
}



# Function to get taxonomy values for a given product type
def get_taxonomy_values(product_type):
    if product_type in taxonomy_dict:
        return ', '.join(taxonomy_dict[product_type])
    else:
        return ''

# Create a new 'taxonomy_column' based on 'Product_Type'
df['taxonomy_column'] = df['Product_Type'].apply(get_taxonomy_values)
df['taxonomy_column']=df['taxonomy_column'].str.replace(',','>')


# Remove anything like "&#9989;" from thedata 
df = df.applymap(lambda x: str(x).replace(r'&#[0-9]+;', '') if isinstance(x, str) else x)





# flag multivariant and single variant:

df['Variant_Flag'] = df.apply(lambda row: 'multi' if row['Relationship_Type'] == 'variation' and row['Parent_Child'] == 'Child' else 'single', axis=1)

### TAGS :
# df['Subcategory'].unique()

split_values = df['taxonomy_column'].str.split('>')

#split_values=split_values.str.replace('\s+', ' ')

df['Category:Name']= split_values.str[0].str.strip()

df['Subcategory'] = np.where(split_values.str[-2].str.strip() == split_values.str[0].str.strip(), 
                             split_values.str[-1].str.strip(), 
                             split_values.str[-2].str.strip())




df['Type']= split_values.str[-1].str.strip()

df['Tags'] ="Category_" + df['Category:Name'] + ', ' + "Subcategory_" + df['Subcategory'] + ', ' + "Type_" + df['Type'] + ', ' + split_values.str[1:-2].apply(lambda x: ', '.join(x)).str.strip(',')

df['Tags']=df['Tags'].str.replace(r'\s+', ' ', regex=True)

#df["ProcType"]= df['Category:Name']

df['Product_Keywords']=df['Generic_Keywords']


                                                ######################  ## Title Defining: ########################
# Function to convert dimensions from Inch to cms
def convert_to_cms(size):
    if pd.notna(size):
        # Remove double quotes from the Size_Name and convert to lowercase
        size = size.lower().replace('"', '')
        
        # Remove 'inch' if it exists
        size = size.split('inch')[0]
        
        # Split the dimensions using 'x'
        dimensions = size.split('x')

        # Convert dimensions from Inch to cms
        dimensions_in_cms = [str(round(float(dim.strip()) * 2.54, 2)) + ' cm' for dim in dimensions]

        # Join the dimensions with ' x ' and return
        return ' x '.join(dimensions_in_cms)
    else:
        return size

# Apply the function to the 'Size_Name' column
df['Dimensions_in_cms'] = df['Size_Name'].apply(convert_to_cms)



# Function to extract dimensions when 'Dimensions_in_cm' and 'Size_Name' are NaN
def extract_dimensions(row):
    if pd.isna(row['Dimensions_in_cms']) and pd.isna(row['Size_Name']):
        if pd.notna(row['Bullet_Point5']):
            start_index = row['Bullet_Point5'].lower().find('or') + 2
            end_index = row['Bullet_Point5'].lower().find('|', start_index)
            extracted_dimensions = row['Bullet_Point5'][start_index:end_index].strip()
            return extracted_dimensions
        elif pd.notna(row['Bullet_Point6']):
            start_index = row['Bullet_Point6'].lower().find('or') + 2
            end_index = row['Bullet_Point6'].lower().find('|', start_index)
            extracted_dimensions = row['Bullet_Point6'][start_index:end_index].strip()
            return extracted_dimensions
    
    return row['Dimensions_in_cms']


# Apply the function to the 'Dimensions_in_cms' column
df['Dimensions_in_cms'] = df.apply(extract_dimensions, axis=1)


def extract_dimensions_again(row):
    if 'cm x' not in row['Dimensions_in_cms'].lower():
        start_index = row['Bullet_Point6'].lower().find('or') + 2
        end_index = row['Bullet_Point6'].lower().find('|', start_index)
        extracted_dimensions = row['Bullet_Point6'][start_index:end_index].strip()
        return extracted_dimensions
    return row['Dimensions_in_cms']



df['Dimensions_in_cms'] = df.apply(extract_dimensions_again, axis=1)


def extract_text(row):
    def extract_for_text(text):
        if pd.notna(text):
            match = re.search(r'FOR\s(.*?):', text, re.IGNORECASE)
            return 'For ' + match.group(1) if match else ''
        return ''

    return extract_for_text(row['Bullet_Point7']) or extract_for_text(row['Bullet_Point5']) or extract_for_text(row['Bullet_Point6'])

    
# Apply the function to the 'Extracted_Text' column
df['Used_for'] = df.apply(extract_text, axis=1)

# df['Used_for'].unique()
indices_blank_used_for = df[df['Used_for'].str.strip().eq('')].index
print(indices_blank_used_for)

# def generate_title(row, taxonomy_dict):
#     product_type = row['Product_Type']
#     parent_child = row['Parent_Child']

#     # Get the corresponding taxonomy for the product type
#     if product_type in taxonomy_dict:
#         categories = taxonomy_dict[product_type][0].split(', ')
#         last_2_categories = ', '.join(categories[-2:])
#         subcategories = last_2_categories if len(categories) >= 2 else last_2_categories

#         # Adjust the title based on parent_child column
#         if parent_child == 'Child':
#             title_parts = [
#                 row['Brand_Name'].strip(),
                
#                 row['Product_Type'].strip(),
#                 row['Used_for'].strip(),
#                 ',',
#                 subcategories.strip(),
#                 ',',
#                 row['Dimensions_in_cms'].strip(),
#                 row['Color_Map'].strip()
#             ]
#         elif parent_child == 'Parent':
#             title_parts = [
#                 row['Brand_Name'].strip(),
#                 row['Product_Type'].strip(),
#                 row['Used_for'].strip(),
#                 ',',
#                 subcategories.strip(),
#                 ', Multisize & Multicolor Available'
#             ]
#         else:
#             raise ValueError("Invalid value in 'parent_child' column.")

#         # Replace NaN values with empty strings and join with spaces
#         title = ' '.join(str(part) if pd.notna(part) else '' for part in title_parts)
#         # Trim multiple spaces from start, middle, and end
#         title = ' '.join(title.split())

#         # Capitalize each word in the title
#         title = title.title().strip()

#         return title
#     else:
#         raise ValueError(f"No taxonomy information found for product type: {product_type}")

# # Assuming 'taxonomy_dict' is defined as mentioned in previous interactions
# # Assuming 'parent_child' column exists in the DataFrame
# # Usage: 
# df['Title'] = df.apply(lambda row: generate_title(row, taxonomy_dict), axis=1)


df['Title'] = df['Brand_Name'] + ' ' + df['Subcategory'] + ' ' + df['Used_for']

df['Title'] = df['Title'].str.title().str.strip()


############################################################################ Specifications:
    
# Function to perform case-insensitive replacement for the entire DataFrame
def replace_case_insensitive(dataframe, old_str, new_str):
    for column in dataframe.columns:
        if dataframe[column].dtype == 'O':  # Only apply to object (string) columns
            mask = dataframe[column].apply(lambda x: pd.isna(x) or not isinstance(x, str) or old_str.lower() not in x.lower())
            dataframe.loc[~mask, column] = dataframe.loc[~mask, column].apply(lambda x: x.replace(old_str, new_str))
    return dataframe

# Apply the function to the entire DataFrame
df = replace_case_insensitive(df, 'IINCLUDES', 'Includes')



# Function to convert inch dimensions to cm
def convert_to_cm(dimensions_str):
    inch_pattern = r'(\d+(\.\d+)?)\s*("|inch\b)'
    matches = re.findall(inch_pattern, dimensions_str)
    
    for match in matches:
        inch_value = float(match[0])
        cm_value = round(inch_value * 2.54, 2)
        dimensions_str = dimensions_str.replace(f"{match[0]}{match[2]}", f"{cm_value} cm")
        
    # Case-insensitive replacements
    dimensions_str = dimensions_str.lower().replace('cmx', 'cm x').replace('inches', '')
    
    return dimensions_str

# Apply the function to specified columns
columns_to_convert = ['Bullet_Point3', 'Bullet_Point5']
for column in columns_to_convert:
    df[column] = df[column].apply(lambda x: convert_to_cm(x) if pd.notna(x) else x)
    
    
    
    
# Specify the columns containing the information
columns_to_combine = ['Bullet_Point2', 'Bullet_Point3', 'Bullet_Point4', 'Bullet_Point5', 'Bullet_Point6', 'Bullet_Point7']

# Combine the specified columns with '<br> •' as separator
df['Specifications'] = df[columns_to_combine].apply(lambda row: '<br> • '.join(row.dropna()), axis=1)

# Capitalize the first letter of each word in the Specifications column
df['Specifications'] = df['Specifications'].apply(lambda x: ' '.join(word.lower() for word in x.split()))

# Add the prefix "<strong>Specifications:</strong>" to the new column
df['Specifications'] = "<strong>Specifications:</strong><br> •" + df['Specifications']

# Remove anything like "&#9989;" from the Specifications column
df['Specifications'] = df['Specifications'].str.replace(r'&#[0-9]+;', '', regex=True)

# Remove specific words from the Specifications column
words_to_remove = ['we', 'our', 'ours', 'us']
pattern = r'\b(?:' + '|'.join(words_to_remove) + r')\b'

df['Specifications'] = df['Specifications'].str.replace(pattern, '', case=False, regex=True)


df['Specifications'] = df['Specifications'].str.replace(pattern, '', case=False)
df['Specifications'] = df['Specifications'].str.replace('|', '<br> • ', case=False)

df['Specifications'] = df['Specifications'].str.replace('Mdf','MDF', case=False)


# Function to remove Product SKU and INcludes from Sepcifications:
def remove_patterns(column):
    pattern = r'<br> •\s*(Includes|Product Sku|Tags):\s*[^<]*'
    return column.str.replace(pattern, '', case=False, regex=True)

df['Specifications']= remove_patterns(df['Specifications'])



def capitalize_after_bullet(text):
    # Use regular expression to find text between "• " and ":"
    matches = re.findall(r'• (.*?):', text)
    
    # Capitalize the first letter of each word
    for match in matches:
        capitalized_match = ' '.join(word.capitalize() for word in match.split())
        text = text.replace(match, capitalized_match, 1)  # Replace the first occurrence
    
    return text

# Apply the function to the 'Specifications' column
df['Specifications'] = df['Specifications'].apply(capitalize_after_bullet)


def extract_features_from_specifications(specifications_series):
    spec_tag = "<strong>Specifications:</strong>"
    feature_tag = "<strong>Features:</strong>"

    # Keywords that might indicate a feature
    feature_keywords = ['design', 'artwork', 'benefit', 'unique', 'advantage', 'functional', 'usage', 'easy to use','precise drawings']

    features_list = []
    specifications_list = []

    for specifications_column in specifications_series:
        # Initialize sections
        specs, features = "", ""

        # Split the description into parts
        parts = re.split(f"({spec_tag})", specifications_column)

        for i, part in enumerate(parts):
            if spec_tag in part:
                specs = parts[i + 1]  # Get the specifications part

        # Check for features in specifications
        if specs:
            # Separate features from specifications if present
            specs_lines = specs.split("<br>")
            specs = ""
            for line in specs_lines:
                if "•" in line and any(keyword in line.lower() for keyword in feature_keywords):
                    features += line + "<br>"
                else:
                    specs += line + "<br>"

            specifications_list.append("<br><strong>Specifications:</strong>" + specs)
            features_list.append("<br><strong>Features:</strong><br>" + features)

    return features_list, specifications_list

# Call the function
features, specifications = extract_features_from_specifications(df['Specifications'])

# Assuming df is your DataFrame
df['Features'] = features
df['Specifications'] = specifications

# Define the regular expression pattern
pattern = re.compile(r'\d+(\.\d+)?\s*cm\s*x\s*\d+(\.\d+)?\s*cm\s*or\s*') #  25.4 cm x 20.32 cm or 25.4 cm x 20.3 cm

# Function to apply the regex and remove the pattern
def remove_pattern(text):
    return re.sub(pattern, '', text)

# Apply the function to the 'Specifications' column
df['Specifications'] = df['Specifications'].apply(remove_pattern)


############################################################################# Package Includes:
import inflect
def create_package(row):
    p = inflect.engine()

    if 'package' in row['Bullet_Point1'].lower():
        # Extract text after 'pack includes' till '.'
        package_text = re.search(r'pack includes(.*?)(\.|$)', row['Bullet_Point1'], re.IGNORECASE)
        package_text = package_text.group(1).strip() if package_text else ''
        package_text = package_text.replace('- Unframed', '').title()
        
        # Find the highest digit in the package_text ignoring numbers before "cm"
        matches = re.findall(r'(\d+)(?!.*cm)', package_text)
        highest_digit = max(map(int, matches), default=1)

        # Replace '1' with the highest digit
        package_text = package_text.replace('1', str(highest_digit), 1)

        # Remove the highest digit from the original place
        package_text = package_text.replace(str(highest_digit), '', 1)

        # Convert the highest_digit to words
        highest_digit_words = p.number_to_words(highest_digit)

        # Format the 'Package' column
        package_column = f'<br><br><strong>Package Includes:</strong><br> • {highest_digit} x {package_text}'

        # Format the 'Pack of' column
        pack_of_column = f'{highest_digit} Piece,'

        return package_column, pack_of_column

    elif 'includes' in row['Bullet_Point1'].lower():
        # Extract text after 'INCLUDES' till ',' or '-'
        package_text = re.search(r'INCLUDES(.*?)(,|-|$)', row['Bullet_Point1'], re.IGNORECASE)
        package_text = package_text.group(1).strip() if package_text else ''
        package_text = package_text.replace('as shown', '').title()
        
        # Find the highest digit in the package_text ignoring numbers before "cm"
        matches = re.findall(r'(\d+)(?!.*cm)', package_text)
        highest_digit = max(map(int, matches), default=1)

        # Replace '1' with the highest digit
        package_text = package_text.replace('1', str(highest_digit), 1)

        # Remove the highest digit from the original place
        package_text = package_text.replace(str(highest_digit), '', 1)

        # Convert the highest_digit to words
        highest_digit_words = p.number_to_words(highest_digit)

        # Format the 'Package' column
        package_column = f'<br><strong>Package Includes:</strong><br> • {highest_digit} x {package_text}'

        # Format the 'Pack of' column
        pack_of_column = f'{highest_digit} Piece,'

        return package_column, pack_of_column

    else:
        # Use 'Title' column as package text
        package_text = row['Title'].split(',')[0]

        # Convert the highest_digit to words
        highest_digit_words = p.number_to_words(1)

        # Format the 'Package' column
        package_column = f'<br><strong>Package Includes:</strong><br> • 1 x {package_text}'

        # Format the 'Pack of' column
        pack_of_column = '1 Piece,'

        return package_column, pack_of_column

# Example usage:
# Assuming df is your DataFrame
df['Package'], df['Pack_of_Column'] = zip(*df.apply(create_package, axis=1))

# Now df['Package_Column'] and df['Pack_of_Column'] contain the formatted package and pack of columns, respectively.

# 1 piece, Inkdotpot Wall Art for Laundry Room, Artwork Decor, Available In Multiple Sizes & Colors
# 1 Piece, Inkdotpot Wall Art for Laundry Room,

# Apply the function to the 'Package' column
#df['Package'] = df.apply(create_package, axis=1)

df['Package'] = df['Package'].str.replace('set of','', case=False)


def replace_text(variation, original_text):
    if variation == 'multi':
        return original_text.replace('<strong>Package Includes:</strong>', '<strong>Package Includes:</strong> <i>(as per your selected option)</i>')
    else:
        return original_text

# Apply the custom function to the DataFrame
df['Package'] = df.apply(lambda row: replace_text(row['Variant_Flag'], row['Package']), axis=1)



##################################################################  Body HTML  #######################################################
df['Body HTML']=  df['Features'] +  df['Specifications'] + df['Package']



# List of stop words
stop_words = set(['CM', ' FOR ', ' of ', 'Cm X', 'and', 'Comes', 'their','-and', ' the ', 'also',' an ', ' you ','as Well', ' is ', ' in ', \
                  ' to ', ' for ', ' with ', 'will', 'X ', ' X ',' your ', 'With', ' on ', ' from ', 'with', ' a ', ' as ', 'kg', 'cm ', ' x ', ' are ', ' so ', 'that',
                   ' it ', 'any', 'Mm'])

# Function to ensure stop words are always in lowercase for specific columns
def ensure_stop_words_lower_specific(column, stop_words):
    for stop_word in stop_words:
        # Use re.IGNORECASE instead of case=False for case-insensitive matching
        pattern = re.compile(re.escape(stop_word), flags=re.IGNORECASE)
        # Replace stop word with lowercase version, regardless of case
        column = column.apply(lambda x: re.sub(pattern, stop_word.lower(), x))
    return column

# Apply the function to specific columns
df['Title'] = ensure_stop_words_lower_specific(df['Title'], stop_words)
df['Body HTML'] = ensure_stop_words_lower_specific(df['Body HTML'], stop_words)
df['Body HTML'] = df['Body HTML'].str.replace('a-z','A-Z', case=False)

def capitalize_after_dot_or_colon(column):
    if pd.notna(column) and isinstance(column, str):
        # Capitalize letters after '.' and ':'
        result = re.sub(r'(?<=[.:])\s*([a-zA-Z])', lambda match: match.group(0).upper(), column)
        return result
    return column
df['Body HTML'] = df['Body HTML'].apply(capitalize_after_dot_or_colon)

def remove_inches(column):
    if pd.notna(column) and isinstance(column, str):
        # Remove patterns like '10.5" x 5.8" Inches Or'
        result = re.sub(r'\d+(\.\d+)?"\s*x\s*\d+(\.\d+)?"\s*inches\s*or', '', column)
        return result
    return column
df['Body HTML'] = df['Body HTML'].apply(remove_inches)


# Assuming 'ID' is the column you want to check for duplicates
duplicate_cases = df[df.duplicated(subset='Variant SKU', keep=False)]

# Set 'body_html' to NaN where 'Parent_Child' is 'Child'
#df.loc[df['Parent_Child'] == 'Child', 'Body HTML'] = np.nan

#df = df.drop_duplicates(subset='ID', keep='first')

######################################################################## Tags ###############################################


# def extract_tags(df):
#     # Create a new column 'Tags'
#     df['Sub_Tag'] = ''

#     # Iterate through rows
#     for index, row in df.iterrows():
#         # Check if 'TAGS:' is present in 'bullet_point6'
#         if 'TAGS:' in str(row['Bullet_Point6']):
#             tags_index = index
#             tags = str(row['Bullet_Point6']).split('TAGS:')[1].strip()#.split(',')[0]
#             df.at[tags_index, 'Sub_Tag'] = tags
#         # Check if 'TAGS:' is present in 'bullet_point8'
#         if 'Tags:' in str(row['Bullet_Point8']):
#             tags_index = index
#             tags = str(row['Bullet_Point8']).split('Tags:')[1].strip().replace('Tags: ','')
#             df.at[tags_index, 'Sub_Tag'] = tags

#     return df

# # Assuming your DataFrame is named 'df'
# df = extract_tags(df)

#Brand_Nife, Category_Evening Dresses, Color_Pink, Gender_Women, not_update_CA, Type_Women's Clothing

# def generate_tags(row):
#     title_parts = [
#         'Brand_' + str(row['Brand_Name']),
#         'Category_' + str(row['Product_Type']),
#         'Color_' + str(row['Color_Map']),
#         'Type_' + str(row['Sub_Tag'])
#     ]

#     # Replace NaN values with empty strings and join with spaces
#     title = ','.join(part for part in title_parts if part.strip())

#     # Trim multiple spaces from start, middle, and end
#     title = ' '.join(title.split())

#     # Capitalize each word in the title
#     title = title.title()

#     return title






# Assuming 'taxonomy_dict' is defined as mentioned in previous interactions
# Usage: df['Tags'] = df.apply(lambda row: generate_tags(row, taxonomy_dict), axis=1)

# df['Tags'] =  df['taxonomy_column'] + "," + df['Sub_Tag']
# df['Tags']=df['Tags'].str.title()
# df['Supplier_Tags']= df['Sub_Tag']#df['Bullet_Point8']
# df['New_Tags']=df['taxonomy_column'].str.replace(',','>')
columns_to_concatenate = [
    'Other_Image_Url1', 'Other_Image_Url2', 'Other_Image_Url3', 
    'Other_Image_Url4', 'Other_Image_Url5', 'Other_Image_Url6', 
    'Other_Image_Url7', 'Other_Image_Url8'
]

df['Image_Column'] = df[columns_to_concatenate].apply(
    lambda row: '|'.join(row.dropna().astype(str)), axis=1
)

###################################################################### Final DataFrame: ######################################

df3_result= df[['Variant SKU','Brand_Name','Title','Body HTML','Category:Name','Type','Tags','Product_Keywords','Parent_Child','Color_Name', 'Color_Map','Quantity','Image_Column']]

# df3_result['Brand_Name'].unique()
# len(df3_result['Title'].unique())
# count= df3_result.groupby('Title')['Variant SKU'].nunique()


def capitalize_after_dot_or_colon(column):
    if pd.notna(column) and isinstance(column, str):
        # Capitalize letters after '.' and ':'
        result = re.sub(r'(?<=[.:•])\s*([a-zA-Z])', lambda match: match.group(0).upper(), column)
        return result
    return column

df3_result['Body HTML'] = df3_result['Body HTML'].apply(capitalize_after_dot_or_colon)


df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('<br><br><br>','<br><br>')


###################################################################### HTML:

output_file_path = initial_raw_file.replace("IBACRAFT_AK", "Traceease_FinalCleanData")

df3_result.to_excel(output_file_path, index=False)

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

 #Calculate the script runtime in minutes
end_time = time.time()
runtime_minutes = (end_time - start_time) / 60

print(f"Script completed in {runtime_minutes:.2f} minutes.")