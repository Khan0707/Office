# -*- coding: utf-8 -*-
"""
Created on Mon Nov 13 08:11:21 2023

@author: aaqib
"""

import pandas as pd
import regex as re
from datetime import datetime
from num2words import num2words


python_script_cleaned = r"D:\BackendData\VidaXL\18_01_Vidaxl_CleanedDesc_NZ_HTML.xlsx"

initial_raw_file= r"D:\BackendData\Vidaxl\18_01_Vidaxl_RawExport_NZ.xlsx"

# Read the Excel file into a DataFrame
df = pd.read_excel(python_script_cleaned, engine='openpyxl')

# Define a pattern to match variations of the character
pattern = re.compile(r'[â€¢]')
# Use regex substitution to replace the character
df['Body HTML'] = df['Body HTML'].apply(lambda x: pattern.sub('•', str(x)))

# Define a regular expression pattern to match consecutive occurrences of •
pattern = re.compile(r'•+')
# Use regex substitution to replace consecutive occurrences with a single •
df['Body HTML'] = df['Body HTML'].apply(lambda x: pattern.sub('•', str(x)))

df2= pd.read_excel(initial_raw_file)



# Function to extract text before ':'
def extract_text_before_colon(input_text):
    specifications_index = input_text.find('<br><br><strong>Specifications:</strong><br>')
    feature_text = input_text[:specifications_index].strip()
    specs_text = input_text[specifications_index:].strip()
    colon_index = feature_text.find(':')
    if colon_index != -1:
        result_text = feature_text[:colon_index].split()[:-2]
        result_text = ' '.join(result_text)
        text = input_text.replace(result_text,'').strip()
        return text
    else:
        return specs_text #input_text[specifications_index:].strip().replace('<br><br>','')

# # Apply the function to the 'description' column
df['Body HTML'] = df['Body HTML'].apply(extract_text_before_colon)



# Extract initial word
df['Initial_Word'] = df['Body HTML'].str.split().str[0]
# Count occurrences
word_counts = df['Initial_Word'].value_counts()
print(f"Starting Words are : {word_counts}")

#---->>>> 
# Function to replace Delivery Conatins with PAckage Includes and add Highligts to the begining of the HTML Column:
   
df['Body HTML'] = '<strong>Features:</strong>' + df['Body HTML']


indexes_without_package_includes = df[~df['Body HTML'].str.contains('Package Includes:', case=False)].index
print(f"Count_indexes_without_package_includes: {len(indexes_without_package_includes)}")
print(f"indexes_without_package_includes: {indexes_without_package_includes}")



package_pattern = {
    'Package included:': '<br><br><strong>Package Includes:</strong>',
    'Packaging included:': '<br><br><strong>Package Includes:</strong>',
    'Packaging includes:': '<br><br><strong>Package Includes:</strong>',
    '<br> Package included: <br>': '<br><br><strong>Package Includes:</strong>',
    'Package list': '<br><br><strong>Package Includes:</strong>',
    'Package List': '<br><br><strong>Package Includes:</strong>',
    'Package Included:': '<br><br><strong>Package Includes:</strong>',
    '<br> Package includes<br>': '<br><br><strong>Package Includes:</strong><br>',
    'Package List:': '<br><br><strong>Package Includes:</strong>',
    '<br> Package Include:': '<br><br><strong>Package Includes:</strong>',
   'Packing List': '<br><br><strong>Package Includes:</strong>',
    'Version Included:': '<br><br><strong>Package Includes:</strong>',
    'Version Includes:': '<br><br><strong>Package Includes:</strong>',
    'Package Contents:': '<br><br><strong>Package Includes:</strong>',
    'Package Content': '<br><br><strong>Package Includes:</strong>',
   '<br> Package included: <br>': '<br><br><strong>Package Includes:</strong><br>',
   'PRODUCT LIST': '<br><br><strong>Package Includes:</strong><br>',
   '<br> Package<br>': '<br><br><strong>Package Includes:</strong><br>',
   'Package included :': '<br><br><strong>Package Includes:</strong><br>',
   '<br> Package included ：<br>': '<br><br><strong>Package Includes:</strong><br>',
   '<br> Package Included: <br> ' : '<br><br><strong>Package Includes:</strong><br>',
   '<br> Pakage Include: <br> ': '<br><br><strong>Package Includes:</strong><br>',
   ("<br> Package Included：<br> "): '<br><br><strong>Package Includes:</strong><br>',
   'Delivery contains:': '<br><br><strong>Package Includes:</strong>'
   , '<br> • Delivery contains: <br>': '<br><br><strong>Package Includes:</strong>'
   , '<br> • Delivery contains : <br>': '<br><br><strong>Package Includes:</strong>'
   , '<br> • Delivery contains': '<br><br><strong>Package Includes:</strong>'
   , '<br> • Delivery includes' :'<br><br><strong>Package Includes:</strong>'
   , 'delivery includes:':'<br><br><strong>Package Includes:</strong>'
   , 'delivering includes' :'<br><br><strong>Package Includes:</strong>'
   , 'The tool set includes:': '<br><br><strong>Package Includes:</strong>'
    
}

df['Body HTML'] = df['Body HTML'].replace(package_pattern, regex=True)


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



indexes_without_package_includes = df[~df['Body HTML'].str.contains('Package Includes:', case=False)].index
print(f"Count_indexes_without_package_includes: {len(indexes_without_package_includes)}")
print(f"indexes_without_package_includes: {indexes_without_package_includes}")



df['Body HTML'] = df['Body HTML'].str.replace("Use a mild soap solutionStoring:".lower(),"Use a mild soap solution.Storing:")

# df['Body HTML'] = df['Body HTML'].str.replace('<br> • Table: <br> • Dimensions:','<br> • Table Dimensions:')
# df['Body HTML'] = df['Body HTML'].str.replace('<br> • Chair: <br> • Dimensions:','<br> • Chair Dimensions:')

# df['Body HTML'] = df['Body HTML'].str.replace('!<br><br><br>','.<br><br>')
# df['Body HTML'].str.contains('delivering includes').sum()


indexes_without_package_includes = df[~df['Body HTML'].str.contains('Package Includes:', case=False)].index
print(f"Count_indexes_without_package_includes: {len(indexes_without_package_includes)}")
print(f"indexes_without_package_includes: {indexes_without_package_includes}")


def move_package_to_end(text):
    start_keyword =  r'<br><br><strong>Package Includes:</strong>'
    end_keyword = r'<br><br><strong>Specifications:</strong><br>'

    # Find the start and end indices of the relevant text
    start_index = text.find(start_keyword)
    end_index = text.find(end_keyword)

    if start_index != -1 and end_index != -1 and start_index < end_index:
        # Extract the relevant text between the keywords, including start_keyword
        included_text = text[start_index:end_index].strip()

        # Remove the extracted text from its original position
        text = text[:start_index] + text[end_index:]
        #Replace commas and periods with '<br> •' in the extracted text
        included_text = re.sub(r'[,.]', '<br> •', included_text)

        # Append the extracted text to the end of the cell
        text = text + included_text

    return text

df['Body HTML'] = df['Body HTML'].apply(move_package_to_end)


def remove_duplicate_substring(df, column_name, substring):
    # Iterate over each row in the specified column
    for index, row in df.iterrows():
        text = row[column_name]

        # Find the first occurrence of the substring
        first_occurrence = text.find(substring)

        # If the substring is found, find the second occurrence
        if first_occurrence != -1:
            second_occurrence = text.find(substring, first_occurrence + 1)

            # If the second occurrence is found, remove everything after it
            if second_occurrence != -1:
                text = text[:second_occurrence]

        # Update the DataFrame with the modified text
        df.at[index, column_name] = text.strip()  # Trim any leading or trailing spaces

    return df

# Call the function to remove duplicate occurrences of the specified substring
df = remove_duplicate_substring(df, 'Body HTML','<br><br><strong>Package Includes:</strong>')





# Merge the DataFrames on 'Variant SKU' and replace 'Body HTML' in df2 with 'Body HTML' in df

df3 = pd.merge(df2, df[['Variant SKU', 'Body HTML']], on='Variant SKU', how='left', suffixes=('_replace', '_original'))
# Use 'Body HTML_original' as the final 'Body HTML' column in df3
df3['Body HTML'] = df3['Body HTML_original'].combine_first(df3['Body HTML_replace'])
# Drop the unnecessary columns
df3 = df3.drop(['Body HTML_replace', 'Body HTML_original'], axis=1)
# Set the column order to match df2
column_order = df2.columns
df3 = df3[column_order]

# Mapping for mattress sizes
mattress_size_mapping = {
    'Single Size': ('92cm x 187cm', '92x187', '90x190'),
    'Single XL Size': ('92cm x 203cm', '92x203', '92x203'),
    'King Single Size': ('106cm x 203cm', '106x203', '107x203'),
    'Double Size': ('137cm x 187cm', '137x187', '135x190'),
    'Queen Size': ('153cm x 203cm', '153x203', '150x200'),
    'King Size': ('183cm x 203cm', '183x203', '183x203'),
    'Super King Size': ('203cm x 203cm', '203x203','203x203')
}

# Mapping for headboard sizes
headboard_size_mapping = {
    'Single Size': ('90 cm', '90','92 cm','92', '91 cm','91'),
    'King Single Size': ('107 cm', '107','106 cm','106','105 cm','105'),
    'Double Size': ('137 cm', '137','135 cm', '135','136 cm', '136'),
    'Queen Size': ('152 cm', '152','153 cm', '153','150 cm', '150'),
    'King Size': ('183 cm', '183','183 cm', '183','183 cm', '183'),
    'Super King Size': ('203 cm', '203','203 cm', '203','203 cm', '203')  # No specific size found
}


# size_mapping = {
#     'Single Size': ['92cm x 187cm', '90 cm'],
#     'Single XL Size': ['92cm x 203cm'],
#     'King Single Size': ['106cm x 203cm', '107 cm'],
#     'Double Size': ['137cm x 187cm', '137 cm'],
#     'Queen Size': ['153cm x 203cm', '152 cm'],
#     'King Size': ['183cm x 203cm', '183 cm'],
#     'Super King Size': ['203cm x 203cm', '203cm']  
# }

from difflib import get_close_matches

def extract_and_append_size(row):
    if 'Subcategory_Bed' in row['Tags'] or 'Subcategory_Mattress' in row['Tags'] or 'Subcategory_Headboard' in row['Tags']:
        # Define a pattern for finding sizes in the Title column
        # size_pattern = re.compile(r'(\d{1,3}\s?[cmxX]\s?\d{1,3}\s?cm|\d{1,3}\s?[cmxX]\s?\d{1,3})')
        size_pattern = re.compile(r'(\d{1,3}(?:\s?[cmxX]\s?\d{1,3})?\s?cm?)', re.IGNORECASE)

        matches = size_pattern.findall(row['Title'])
        
        
        if matches:
            # Extract the first size found in the Title
            
            size = matches[0].lower()
            
            
            # if 'Subcategory_Bed' in row['Tags'] or 'Subcategory_Mattress' in row['Tags']:
            #     size_mapping = mattress_size_mapping
            # elif 'Subcategory_Headboard' in row['Tags']:
            #     size_mapping = headboard_size_mapping
            
            if 'Subcategory_Bed' in row['Tags'] or 'Subcategory_Mattress' in row['Tags']:
                size_mapping = mattress_size_mapping
                size_match = re.search(r'\d+x\d+', size)
                if size_match:
                    size = size_match.group()
                closest_match = get_close_matches(size, [value[0] for value in size_mapping.values()] 
                                                      + [value[1] for value in size_mapping.values()]
                                                      + [value[2] for value in size_mapping.values()], n=1, cutoff=0.8)

            elif 'Subcategory_Headboard' in row['Tags']:
                # For headboards, extract only the second number (90 in "112X90 cm")
                size_parts = re.findall(r'\d+', size)
                if size_parts:
                    size = size_parts[-1]
                size_mapping = headboard_size_mapping
                
                closest_match = get_close_matches(size, [value[0] for value in size_mapping.values()] 
                                                      + [value[1] for value in size_mapping.values()]
                                                      + [value[2] for value in size_mapping.values()] 
                                                      + [value[3] for value in size_mapping.values()]
                                                      + [value[4] for value in size_mapping.values()]
                                                      + [value[5] for value in size_mapping.values()], n=1, cutoff=0.8)

          
            
            if closest_match:
                closest_size_value = closest_match[0]
                
                # Find the corresponding key for the closest size value
                corresponding_key = next((key for key, value in size_mapping.items() if closest_size_value in value), '')
                
                if corresponding_key:
                    row['Title'] = f"{corresponding_key} {row['Title']}"
    
    return row


df3 = df3.apply(extract_and_append_size, axis=1)



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

# test=df3_result.loc[df3_result['Variant SKU']=='PIMW-94028']

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
columns_to_convert = ['ID', 'Variant Inventory Item ID', 'Variant ID', 'Variant Barcode']
df3_result[columns_to_convert] =df3_result[columns_to_convert].astype(str)


## Adding Package Includes for those were it is missing:
def add_package_includes(row):
    package_text = row['Title']

    if package_text and package_text[0].isdigit():
        # Convert the digit to its English word equivalent
        digit_word = num2words(int(package_text[0])).title()
        
        # Replace the digit with the word in the text
        package_text = digit_word + package_text[1:]
    
    if '<strong>package includes:</strong>' not in row['Body HTML'].lower() and 'units' in row['Title'].lower():
        
        match = re.search(r'(\d+)\s*Units', package_text) # Extract Quantity
        
        if match:
            extracted_number = str(int(match.group(1)))
        else:
            extracted_number= '000'
        
        text_title = re.split(r'(\d)', package_text, maxsplit=1)[0].strip() # Extract Text Content

        modified_package_text = text_title
        highest_digit = extracted_number

        # Format the 'Package' column
        package_column = f'<br><br><strong>Package Includes:</strong><br> • {highest_digit} x {modified_package_text}'

        return row['Body HTML'] + package_column
    
    elif '<br><br><strong>Package Includes:</strong>' not in row['Body HTML']:
        
        text_title = re.split(r'(\d)', package_text, maxsplit=1)[0].strip()
        
        package_includes_str = '<br><strong>Package Includes:</strong><br> • 1 x '
        
        row['Body HTML'] += package_includes_str + text_title
        
        return row['Body HTML']
    else:
        return row['Body HTML']
 
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


# Capitalise lines of Highlights section:
def capitalize_sentences(text):
    start_pattern = r'<strong>Features:</strong>'
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

# Perform the specified replacements
def replace_before_colon(text):
    start_pattern = r'<strong>Features:</strong>'
    end_pattern = r'<br><br><strong>Specifications:</strong><br>'

    # Find the start and end positions
    start_pos = re.search(start_pattern, text)
    end_pos = re.search(end_pattern, text)

    if start_pos and end_pos:
        start_pos = start_pos.end()
        end_pos = end_pos.start()

        # Extract the text between the specified patterns
        highlighted_text = text[start_pos:end_pos]

        # Replace the word before colon with '<br> • ' before the word and capitalize the word after colon
        highlighted_text = re.sub(r'(\b\w+)\s*:', r'<br> • \1:', highlighted_text)

        # Replace the original highlighted text with the modified version
        text = text[:start_pos] + highlighted_text + text[end_pos:]

    return text

df3_result['Body HTML'] = df3_result.apply(add_package_includes, axis=1)  


df3_result['Body HTML'] = df3_result['Body HTML'].apply(clean_package_includes)

indexes_without_package_includes = df3_result[~df3_result['Body HTML'].str.contains('Package Includes:', case=False)].index
print(f"Count_indexes_without_package_includes: {len(indexes_without_package_includes)}")
print(f"indexes_without_package_includes: {indexes_without_package_includes}")



def syntax_package_includes(sentence):
    # Check if the sentence contains "<strong>Package Includes:</strong>"
    if "<strong>Package Includes:</strong>" in sentence:
        # Extract the content after "<strong>Package Includes:</strong>"
        package_content = sentence.split("<strong>Package Includes:</strong>")[-1].strip()
        
        # Split the content into items
        items = package_content.split(' and ')
        
        # Format each item
        formatted_items = [f"<br>• {item.strip()}" for item in items]
        
        # Join the formatted items into a single string
        result = sentence.split("<strong>Package Includes:</strong>")[0] + '<strong>Package Includes:</strong>'+ ''.join(formatted_items)
        
        return result

    # Return the original sentence if it doesn't contain the specified pattern
    return sentence

df3_result['Body HTML'] = df3_result['Body HTML'].apply(syntax_package_includes)


df3_result['Body HTML'] =df3_result['Body HTML'].str.replace('Note:', '')



df3_result['Body HTML'] =df3_result['Body HTML'].str.replace('Good to know:', '<br><br> • <strong> Note: </strong>') 


###------>>>> df3_result['Body HTML'] .str.contains('Use a Mild Soap Solutionstoring').sum()

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
        patterns_between_br = ['<br>•<br>', '<br>• <br>','<br> •<br>','<br> • <br>']
        for pattern in patterns_between_br:
            while pattern in text:
                text = text.replace(pattern, '<br>')
        
        # Remove all occurrences of <br>• at the end
        text = re.sub(r'<br>•\s*$', '', text)
    
    return text

df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_br_before_strong_and_between_br)



def uppercase_after_punctuation(input_string):
    pattern = re.compile(r'(?<=[.:]|<br><br> • <strong> Note: </strong>)\s*([a-zA-Z])')
    result_string = pattern.sub(lambda x: x.group(0).upper(), input_string)
    return result_string

df3_result['Body HTML'] = df3_result['Body HTML'].apply(uppercase_after_punctuation)


def replace_exclamation(row):
    pattern1 = r'<strong>Features:</strong>'
    pattern2 = r'<br><br><strong>Specifications:</strong><br>'
    text = row['Body HTML']
    sku_variant = row['Variant SKU']

    # Original replacement logic
    regex_pattern = re.compile(f'{pattern1}(.*?){pattern2}', re.DOTALL)

    matches = regex_pattern.finditer(text)

    for match in matches:
        if match.group(1).rstrip().endswith('!'):
            updated_text = text[:match.start(1)] + match.group(1).rstrip()[:-1] + '.' + text[match.end(1):]
            print(f"Replacement occurred in SKU Variant: {sku_variant}")
            return updated_text

    return text

# Apply the function to the 'Body HTML' column
df3_result['Body HTML'] = df3_result.apply(replace_exclamation, axis=1)


# Highlights in Bullets: 
def extract_and_append(row):
    start_pattern = r'<strong>Features:</strong>'
    end_pattern = r'<br><br><strong>Specifications:</strong><br>'

    # Extract text between start and end patterns
    match = re.search(f'{start_pattern}(.*?){end_pattern}', row, re.DOTALL)
    if match:
        highlighted_text = match.group(1)

        # Check if there is no occurrence of '!' and ':'
        if '!' not in highlighted_text and ':' not in highlighted_text :
            # Replace '. ' with '.<br>'
            return row.replace('. ', '.<br>')
        split_highlighted_text = highlighted_text.split('!')

        # Check if the split resulted in at least two elements
        if len(split_highlighted_text) > 1:
            highlighted_text1 = split_highlighted_text[1]

            # Apply the extraction and formatting to the highlighted text
            sentences = highlighted_text.split('!')
            extracted_words = []

            for sentence in sentences:
                parts = sentence.split('.')
                for part in parts:
                    if ':' in part:
                        sub_parts = part.split(':')
                        extracted_word = sub_parts[0].strip()
                        following_part = sub_parts[1].strip().split('>')
                        following_word = '>'.join([word.strip().capitalize() for word in following_part])
                        extracted_words.append('<br> • ' + extracted_word + ': ' + following_word)

            return row.replace(highlighted_text1, ' '.join(extracted_words))
        else:
            sentences = highlighted_text.split('!')
            extracted_words = []

            for sentence in sentences:
                parts = sentence.split('.')
                for part in parts:
                    if ':' in part:
                        sub_parts = part.split(':')
                        extracted_word = sub_parts[0].strip()
                        following_part = sub_parts[1].strip().split('>')
                        following_word = '>'.join([word.strip().capitalize() for word in following_part])
                        extracted_words.append('<br> • ' + extracted_word + ': ' + following_word)

            return row.replace(highlighted_text, ' '.join(extracted_words))
    else:
        return row

df3_result['Body HTML'] = df3_result['Body HTML'].apply(extract_and_append)


def remove_empty_highlights(row):
    start_pattern = r'<strong>Features:</strong>'
    end_pattern = r'<br><br><strong>Specifications:</strong><br>'
    
    # Extract text between start and end patterns (case-insensitive)
    match = re.search(f'{start_pattern}(.*?){end_pattern}', row, re.DOTALL | re.IGNORECASE)
    if match:
        highlighted_text = match.group(1)
        
        # Check if there is only '<br><br>' and no other text
        if highlighted_text.strip() == '':
            # Remove the entire block (case-insensitive)
            return re.sub(f'{start_pattern}.*?{end_pattern}', '<strong>Specifications:</strong><br>', row, flags=re.DOTALL | re.IGNORECASE)
    return row

#df3_result['Body HTML'].apply(lambda x: len(x) > 1600).sum()

#df3_result['Body HTML'].str.contains('!<br><br><br>').sum()

def remove_large_highlights(row):
    start_pattern = r'<strong>Features:</strong>'
    end_pattern = r'<br><br><strong>Specifications:</strong><br>'
    
    # Extract text between start and end patterns
    match = re.search(f'{start_pattern}(.*?){end_pattern}', row, re.DOTALL)
    if match:
        highlighted_text = match.group(1)
        
        # Check if there is only '<br><br>' and no other text, and if character count exceeds 1600
        if len(row) > 1600:
            # Remove the entire block
            return row.replace(f'{start_pattern}{highlighted_text}{end_pattern}','<strong>Specifications:</strong><br>')
    return row

df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_empty_highlights)

df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_large_highlights)

df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_br_before_strong_and_between_br)

def convert_after_br(input_text):
    # Define a custom list of stop words
    custom_stop_words = set(['and', '-and','the', 'an', 'of', 'is', 'in', 'to', 'for', 'with', 'on','from','with','a','as','kg', 'cm', 'x', 'are','so','that'])

    # Split the text based on '<br>•'
    parts = input_text.split('<br> •')

    # Iterate through the parts starting from the second onek
    for i in range(1, len(parts)):
        # Convert the text after each '<br>•' to title case, excluding custom stop words
        words = parts[i].strip().split()
        title_case_words = [word.title() if word.lower() not in custom_stop_words else word for word in words]
        parts[i] = ' '.join(title_case_words)

    # Join the parts back together
    result_text = '<br>• '.join(parts)

    return result_text

df3_result['Body HTML'] = df3_result['Body HTML'].apply(convert_after_br)

#################### TO BE TESTED #############
# Replace "High Gloss" with "Glossy Look" in 'Title' and 'Body HTML'
df3_result['Title'] = df3_result['Title'].str.replace('High Gloss', 'Glossy Look')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('High Gloss', 'Glossy Look')

def move_line_to_end(text):
    start_keyword = 'Specifications'
    end_keyword = '<Strong>Package Includes:'
    target_text = "This Product Packaged and Delivers in"

    # Find the start and end indices of the relevant text
    start_index = text.find(start_keyword)
    end_index = text.find(end_keyword)

    if start_index != -1 and end_index != -1 and start_index < end_index:
        # Extract the relevant text between the keywords
        specifications_text = text[start_index + len(start_keyword):end_index].strip()

        # Find the line containing the target text
        lines = specifications_text.split('<br>• ')
        for i in range(1, len(lines)):
            if target_text in lines[i]:
                # Move the matching line to the end
                lines.append(lines.pop(i))
                break

        # Replace the original text with the modified text
        text = text[:start_index + len(start_keyword)] + '<br>• '.join(lines) + text[end_index:]

    return text

df3_result['Body HTML'] = df3_result['Body HTML'].apply(move_line_to_end)


df3_result['Body HTML'] = df3_result['Body HTML'].replace('<Strong>Package Includes:</Strong>', '<br><br><strong>Package Includes:</strong>', regex=True)


df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('<Br> <Br><br>•','<br>•')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('<Br><Br><br>•','<br>•')

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('!','!<br>')



df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(r'\.(?!\s*(?:max\.|min\.))', '.<br>', flags=re.IGNORECASE)

patterns_to_replace = ['Please note', 'Important note']
replacement = 'Note'

for pattern in patterns_to_replace:
    df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(pattern, replacement, flags=re.IGNORECASE)


def remove_note_string(row):
    start_pattern = r'<strong>Features:</strong>'
    end_pattern = r'<Br><Br><Strong>Specifications:</Strong>'
    note_pattern = r'<br>• Note:.*?$'
    
    # Extract text between start and end patterns
    match = re.search(f'{start_pattern}(.*?){end_pattern}', row, re.DOTALL)
    
    if match:
        highlighted_text = match.group(1)
        
        # Remove the string containing "<br>• Note:" and text after it until the end_pattern
        highlighted_text_cleaned = re.sub(note_pattern, '', highlighted_text, flags=re.MULTILINE)
        
        # Replace the original highlighted text with the cleaned version
        return row.replace(match.group(1), highlighted_text_cleaned)
    else:
        return row

# Apply the function to the 'Body HTML' column
df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_note_string)

df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_empty_highlights)


df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(r'<br>• \d{1}\.', '<br>• ', regex=True)

def remove_duplicate_warning(text):
    # Find the data between '<Strong>Specifications:</Strong>' and '<Br><Strong>Package Includes:' (case-insensitive)
    specs_to_package_match = re.search(r'<Strong>Specifications:</Strong>(.*?)<Br><Strong>Package Includes:', text, flags=re.IGNORECASE)
    
    # If the match is found, check for duplicate occurrences of '<br>• Warning:'
    if specs_to_package_match:
        specs_data = specs_to_package_match.group(1)
        # Count the occurrences of '<br>• Warning:'
        warning_count = specs_data.lower().count('<br>• warning:')
        
        # If there are two or more occurrences, remove lines containing '<br>• Warning:'
        if warning_count >= 2:
            text = re.sub(r'<br>• Warning:.*?<Br><Strong>Package Includes:', '', text, flags=re.IGNORECASE | re.DOTALL)

    return text

df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_duplicate_warning)
#df3_result['Body HTML'].iloc[0:2] = df3_result['Body HTML'].iloc[0:2].str.replace('<br>• 1 x', '<br><br><strong>Package Includes:</strong><br> • 1 x')

clean="<br>• Clean: Use a Mild Soap Solution<br>• Storing: If Possible, Store in a Cool, Dry Place Indoors"
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(clean,'', flags=re.IGNORECASE)


df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_empty_highlights)



# Remove duplicate text lines
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(r'(<br>• .+?)(?=\1)', '', regex=True)

# Reformatt Highlight section:
def reformat_highlights(text):
    start_pattern = r'<strong>Features:</strong>'
    end_pattern = r'<Br><Br><Strong>Specifications:</Strong>'
    # Find the start and end positions
    start_pos = re.search(start_pattern, text, flags=re.IGNORECASE)
    end_pos = re.search(end_pattern, text, flags=re.IGNORECASE)

    if start_pos and end_pos:
        start_pos = start_pos.end()
        end_pos = end_pos.start()

        # Extract the text between the specified patterns
        highlighted_text = text[start_pos:end_pos]

        # Check if '<br>•' exists between the two patterns
        if '<br>•' not in highlighted_text:
            # If not, replace '.' and ',' with '<br>•'
            highlighted_text = re.sub(r'<br>', '', highlighted_text, flags=re.IGNORECASE)
            highlighted_text = re.sub(r'[.]', '<br>• ', highlighted_text, flags=re.IGNORECASE)

            # Replace the original highlighted text with the modified version
            text = text[:start_pos] + highlighted_text + text[end_pos:]

    return text

df3_result['Body HTML'] = df3_result['Body HTML'].apply( reformat_highlights)

df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_br_before_strong_and_between_br)

#df3_result['Body HTML'].str.contains("<strong>Package Includes:</strong>").sum()
df3_result['Body HTML']=df3_result['Body HTML'].str.replace("<br><br><strong>Package Includes:</strong><br><br><strong>Package Includes:</strong><br>",'<br><br><strong>Package Includes:</strong><br>')

indexes_without_package_includes = df3_result[~df3_result['Body HTML'].str.contains('Package Includes:', case=False)].index
print(f"Count_indexes_without_package_includes: {len(indexes_without_package_includes)}")
print(f"indexes_without_package_includes: {indexes_without_package_includes}")

import re 
# MAnual Replcement:
df3_result['Body HTML']=df3_result['Body HTML'].str.replace("<Br>Loading","Loading")
df3_result['Body HTML']=df3_result['Body HTML'].str.replace("<Br> <br>• This Product","<br>• This Product")
df3_result['Body HTML']=df3_result['Body HTML'].str.replace(r'1 Brown Bar Table and 4 Bar Chairs',"<br>• 1 x Brown Bar Table <br>• 4 x Bar Chairs")
df3_result['Body HTML']=df3_result['Body HTML'].str.replace(r"1 Table and 2 Stools","<br>• 1 x Table <br>• 2 x Stools")

df3_result['Body HTML']=df3_result['Body HTML'].str.replace(r'1 Tractor Bar Table and 4 Bar Chairs','<br>• 1 x Tractor Bar Table <br>• 4 x Bar Chairs')

df3_result['Body HTML']=df3_result['Body HTML'].str.replace(r"1 Bar Table and 4 Bar Stools","<br>• 1 x Bar Table <br>• 4 x Bar Stools")
df3_result['Body HTML']=df3_result['Body HTML'].str.replace(r"WITH","with")
df3_result['Body HTML']=df3_result['Body HTML'].str.replace('<Br>• •','<Br>•')
strings_to_replace = ['• Fabric: Polyester: 100%', '• Assembly is Quite Easy','• Max.Loading Capacity: 110 Kg','• Easy Assembly','• Fabric: Leather: 100%','<br>• Legal Documents: More Details About Preventing Your Furniture from Tipping Over Can Be Found']



# Define the keywords
keyword1 =  r'<strong>Features:</strong>'
keyword2 = r'<Strong>Specifications:</Strong>'


# Replace text between keywords
#df3_result['Body HTML'].iloc[0] = re.sub(f'{keyword1}.*?{keyword2}', f'{keyword2}',df3_result['Body HTML'].iloc[0], flags=re.DOTALL)
#df3_result['Body HTML'].iloc[0]=df3_result['Body HTML'].iloc[0].replace(r'<br>• Supplied in a Blow-Moulded Case<Br>','')

#df3_result['Body HTML'].iloc[203]=df3_result['Body HTML'].iloc[203].replace(r'a Bed Frame Only<br>• the Mattress is Not Included<br>• You Can Check Our Shop for the Matching Mattresses','1 x Bed Frame')
                                                                        
                                                                        
#df3_result['Body HTML'].str.contains('<br>• Supplied in a Blow-Moulded Case<Br>').sum()

for string in strings_to_replace:
    df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(string, '')


# Add '<br>' to 'Body HTML' for instances where it doesn't end with '<br>'
df3_result.loc[~df3_result['Body HTML'].str.endswith('<Br>'), 'Body HTML'] += '<Br>'


def remove_spaces(text):
    # Remove spaces before ':', ',', 'mm', 'cm', 'kg', '('
    text = re.sub(r'\s+(:|,|mm|cm|kg|\()', r'\1', text, flags=re.IGNORECASE)
     
    # # Remove spaces within expressions like 'w x h x d'
    # text = re.sub(r'\s*([wWhHdDxX])\s*([wWhHdDxX])\s*([wWhHdDxX])\s*', r'\1x\2x\3', text, flags=re.IGNORECASE)
    
    return text

# Assuming 'Body HTML' is the column you want to process
df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_spaces)



output_file_path = initial_raw_file.replace("RawExport", "FinalCleanData")

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