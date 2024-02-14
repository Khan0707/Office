# -*- coding: utf-8 -*-
"""
Created on Wed Dec 20 08:24:10 2023

@author: aaqib
"""

import pandas as pd
import regex as re
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")#, category=DeprecationWarning)
import time
import numpy as np

python_script_cleaned_base = r"D:\BackendData\Bdit\20_12_BDIT_CleanedDesc_xyz_HTML.xlsx"
initial_raw_file_base= r"D:\BackendData\Bdit\20_12_BDIT_RawExport_xyz.xlsx"

# # List of countries
countries = ['AU','NZ','GKZ'] 


for country in countries:
    #country=""
    print(country)
    python_script_cleaned = python_script_cleaned_base.replace('xyz', country)
    initial_raw_file = initial_raw_file_base.replace('xyz', country)
    
    # Read the Excel file into a DataFrame
    df = pd.read_excel(python_script_cleaned)
    df2= pd.read_excel(initial_raw_file)
    df2 = df2[df2['Body HTML'].notna()]
    
    df['Body HTML'] = '<br><br><strong>Specifications:</strong><br><br> ' + df['Body HTML']
    
    ################# **** Titles Unique String ***** ##################
    
    if country in ['AU', 'NZ']:
        # Function to extract new identifier from SKU
        def extract_identifier(sku):
            try:
                # Split by underscore and take the first part
                first_part = sku.split('_')[0]
                # Split by dash and take the last part
                last_part = sku.split('-')[-1]
                # Extract numbers from the last part (assuming the identifier is numeric)
                numeric_last_part = ''.join(filter(str.isdigit, last_part))
                # Return the new identifier
                return f"{first_part}_{numeric_last_part}"
            except Exception:
                return None  # Return None if there's any issue with processing
    
        # Apply the function to create a new column with the extracted identifiers
        df['New Identifier'] = df['Variant SKU'].apply(extract_identifier)
    
        # Extract unique identifier from "Variant SKU" column
        df['Output'] = df['Variant SKU'].str.extract(r'([A-Za-z0-9_]+)_([A-Za-z0-9]+)', expand=False).apply(lambda x: f"{x[0]}_{x[1]}" if pd.notnull(x[1]) and len(x[1]) <= 7 else x[0], axis=1)
    
        df['Final Output'] = np.where(df['Output'].duplicated(), df['New Identifier'], df['Output'])
    
        # Function to process each value
        def process_value(value):
            if '-' in value and '_' in value:
                last_dash_index = value.rindex('-')
                last_underscore_index = value.rindex('_')
    
                # Choose the delimiter which comes last
                delimiter = '-' if last_dash_index > last_underscore_index else '_'
            elif '-' in value:
                delimiter = '-'
            elif '_' in value:
                delimiter = '_'
            else:
                delimiter = None
    
            if delimiter:
                parts = value.split(delimiter)
                processed_parts = [str(part)[:2] for part in parts]
                return ''.join(processed_parts).upper()
            else:
                return value.upper()
    
        # Apply the function to 'Final Output' column
        df['Final Output'] = df['Final Output'].apply(process_value)
    
    elif country == 'GKZ':
        def map_numbers_to_letters(number):
            mapping = {'0': 'A', '1': 'B', '2': 'C', '3': 'D', '4': 'E', '5': 'F', '6': 'G', '7': 'H', '8': 'I', '9': 'J'}
            result = ''.join(mapping[digit] for digit in str(number))
            return result
    
        # Assuming df is your DataFrame
        df['Final Output'] = df['Variant SKU'].str.split('-').apply(lambda parts: map_numbers_to_letters(parts[1]) if len(parts) > 1 else None)
    
    
    
    # Find instances with more than one occurrence in 'Final Output1'
    duplicates = df[df['Final Output'].duplicated(keep=False)]
    
    # Print the identified instances
    for value in duplicates['Final Output'].unique():
        print("Duplicate Title Strings:")
        print(value)
    
    
    ################# **** Merge ***** ##################
    
    # Merge the DataFrames on 'Variant SKU' and replace 'Body HTML' in df2 with 'Body HTML' in df
    df3 = pd.merge(df2, df[['Variant SKU', 'Body HTML', 'Final Output']], on='Variant SKU', how='left', suffixes=('_replace', '_original'))
    
    # Use 'Body HTML_original' as the final 'Body HTML' column in df3
    df3['Body HTML'] = df3['Body HTML_original'].combine_first(df3['Body HTML_replace'])
    # Drop the unnecessary columns
    df3 = df3.drop(['Body HTML_replace', 'Body HTML_original'], axis=1)
    # Set the column order to match df2
    #column_order = df2.columns
    #df3 = df3[column_order]
    
    
    
    # Step 3: Replace values in other columns
    df3['Status'] = df3['Status'].replace({'Draft': 'Active'})
    df3['Published'] = df3['Published'].replace({False: True})
    df3['Published Scope'] = df3['Published Scope'].replace({'web': 'global'})
    
    # Convert a set of columns from numeric to text format
    columns_to_convert = ['ID', 'Variant Inventory Item ID', 'Variant ID', 'Variant Barcode']
    df3[columns_to_convert] =df3[columns_to_convert].astype(str)
    
    
    ################# **** Tags ***** ##################
        
    # Split the Tags column into separate columns
    df3 = df3.replace(to_replace=r'(?i)color', value='Colour', regex=True)


    df_tags = df3['Tags'].str.split(', ', expand=True)
    
    # Create a dictionary to store tag values
    tag_dict = {}
    
    # Populate the dictionary with tag values
    for col in df_tags.columns:
        for index, values in enumerate(df_tags[col].str.split('_')):
            tag_name = values[0] if isinstance(values, list) and len(values) > 0 else None
            tag_value = values[1] if isinstance(values, list) and len(values) > 1 else None
            if tag_name is not None:
                if tag_name not in tag_dict:
                    tag_dict[tag_name] = [None] * index  # Fill with None until current index
                tag_dict[tag_name].append(tag_value)
    
    # Fill any remaining missing values with None
    max_length = max(len(values) for values in tag_dict.values())
    for tag_name, values in tag_dict.items():
        tag_dict[tag_name] += [None] * (max_length - len(values))
    
    df3 = pd.concat([df3, pd.DataFrame(tag_dict)], axis=1)
    # Drop rows where 'ID' is NaN
    df3 = df3.dropna(subset=['ID'])
    
    df3 = df3.drop(columns=[col for col in df3.columns if col.lower() == 'out of stock'])
    
    
    ################# **** Features: ***** ##################
    
    df3['Features'] = (
    "<br><strong>Features:</strong><br>"
    + "<br> • Brand: " + df3['Brand'].str.title().fillna('') +
    "<br> • Gender: " + df3['Gender'].str.title().fillna('') +
    "<br> • Colour: " + df3['Colour'].str.title().fillna('')
                    )

    ################# **** Titles Making ***** ##################
    df3['Title'] = (
        df3['Brand'].str.title().fillna('') + ' ' +
        df3['Final Output'].fillna('') + ' ' +
        df3['Subcategory'].str.title().fillna('') + ' for ' +
        df3['Gender'].str.title().fillna('') + ' ' +
        df3['Colour'].str.title().fillna('')
    )
    
    df3['Title'] = df3['Title'].str.strip()
    
    df3['Handle'] = df3['Title'].str.lower().str.replace(' ', '-')
    
    ################# **** Package Includes ***** ##################
    package_includes_str = '<br><strong>Package Includes:</strong><br> • 1 x '    
    
    df3['Package'] = package_includes_str + df3['Title'] 
    
    ################# **** Body HTML ***** ##################
    
    df3['Body HTML']= df3['Features']+ df3['Body HTML']+df3['Package'] +'<br>'
    
    # Set the column order to match df2
    column_order = df2.columns
    df3 = df3[column_order]
    
    # Add Bullets to all three sections:  If the substring is not found, it returns -1.
    
    def replace_br_between_keywords(text):
        if pd.notna(text):
            # highlight_index = text.find("<strong>Highlights:</strong>")
            # if highlight_index != -1:
            # # Find the starting index of "<strong>Features: </strong>"
            features_index = text.find("<strong>Features:</strong>")
            if features_index != -1:
                # Find the starting index of "<strong>Specifications:</strong>"
                specs_index = text.find("<strong>Specifications:</strong><br>")
                if specs_index != -1:
                    # Find the starting index of "<strong>Package Includes:</strong>"
                    package_index = text.find("<strong>Package Includes:</strong>")
                    if package_index != -1:
                        # Remove all occurrences of • between the specified keywords
                        #highlights_specs= text[highlight_index:features_index]
                        features_to_specs = text[features_index:specs_index].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>')
                        specs_to_package = text[specs_index:package_index].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>')
                        package_to_end = text[package_index:].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>')
                        # Replace <br> and <br> • with <br> • between the specified keywords
                        #text = text[features_index:specs_index].replace('<br> ', '<br>  •').replace('<br> •', '<br>  •') + text[specs_index:package_index].replace('<br>', '<br>  •').replace('<br> •', '<br>  •') + text[package_index:].replace('<br>', '<br>  •').replace('<br> •', '<br>  •')
                        text = (#highlights_specs.replace('.', '.<br>') +
                            features_to_specs.replace('<br>', '<br>•') +
                            specs_to_package.replace('<br>', '<br>•') +
                            package_to_end.replace('<br>', '<br>•')
                            )
    
        return text
    
    df3['Body HTML'] = df3['Body HTML'].apply(replace_br_between_keywords)
    
    # Perform the specified replacements
    def replace_before_colon(text):
        start_pattern = r'<strong>Features:</strong><br>'
        end_pattern = r'<br><strong>Package Includes:</strong><br>'
    
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
    df3['Body HTML'] = df3['Body HTML'].apply(replace_before_colon)
    
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
    df3['Body HTML'] = df3['Body HTML'].apply(remove_br_before_strong_and_between_br)
    
    def remove_spaces(text):
        # Remove spaces before ':', ',', 'mm', 'cm', 'kg', '('
        text = re.sub(r'\s+(:|,|mm|cm|kg|\()', r'\1', text, flags=re.IGNORECASE)
        
        # Remove spaces within expressions like 'w x h x d'
        #text = re.sub(r'\s*([wWhHdDxX])\s*([wWhHdDxX])\s*([wWhHdDxX])\s*', r'\1x\2x\3', text, flags=re.IGNORECASE)
        
        return text
    df3['Body HTML'] = df3['Body HTML'].apply(remove_spaces)
    
    # Capitalise lines of Highlights section:
    def capitalize_sentences(text):
        start_pattern = r'<strong>Features:</strong>'
        end_pattern = r'<br><strong>Package Includes:</strong><br>'
        # Define a custom list of stop words
        custom_stop_words = set(['and', '-and', 'the', 'an', 'of', 'is', 'in', 'to', 'for', 'with', 'on', 'from', 'with', 'a', 'as', 'kg', 'cm', 'x', 'are', 'so', 'that'])
    
        # Find the start and end positions
        start_pos = re.search(start_pattern, text)
        end_pos = re.search(end_pattern, text)
    
        if start_pos and end_pos:
            start_pos = start_pos.end()
            end_pos = end_pos.start()
    
            # Extract the text between the specified patterns
            highlighted_text = text[start_pos:end_pos]
    
            # Capitalize the first letter of each sentence, excluding stop words
            sentences = re.split(r'(?<=[.!?])\s*', highlighted_text)
            final_text = [sentence.strip().title() if sentence.lower().strip() not in custom_stop_words else sentence.strip() for sentence in sentences]
            capitalized_text = '. '.join(final_text)
    
            # Remove extra full stops after punctuation signs
            capitalized_text = re.sub(r'(?<=[.!?])\s*\.', '', capitalized_text)
    
            # Replace the original highlighted text with the modified version
            text = text[:start_pos] + capitalized_text + text[end_pos:]
    
        return text
    
    df3['Body HTML'] = df3['Body HTML'].apply(capitalize_sentences)
    
    def make_tags_lowercase(text):
        patterns = [r'<Br>', r'<Strong>', r'</Strong>']
    
        for pattern in patterns:
            text = re.sub(re.escape(pattern), pattern.lower(), text)
    
        return text
    
    df3['Body HTML'] = df3['Body HTML'].apply(make_tags_lowercase)
    
    remove_patterns2 = [
       r'<br>• Out Of Stock', r'<br>• New Arrivals',  r'<br>• Gender: Man',  r'<br>• Gender: Woman',  r'<br>• Collection: Fall/Winter',
       r'<br>• External Pockets: Undefined', r'<br>• Internal Pockets: Undefined', r'<br>• Pockets: Undefined',
       r'<br>• Model Wears A Size: Undefined', r'<br>• Platform Height in cm: Undefined'
       
    ]
    
    # Replace additional patterns
    for pattern in remove_patterns2:
        df3['Body HTML']  = df3['Body HTML'] .str.replace(pattern, '', regex=True, flags=re.IGNORECASE)
    
    # Space between Decimals
    df3['Body HTML'] = df3['Body HTML'].apply(lambda x: re.sub(r'(\d+)\s*\.\s*(\d+)', r'\1.\2', x))
    
    df3['Body HTML']  = df3['Body HTML'].str.replace('Heightcm','Height in cm')
    df3['Body HTML']  = df3['Body HTML'].str.replace(',Cm','cm')
    
    
    # Remove small Specs:

    def replace_text_between_patterns(text):
        # Define start and end patterns
        start_pattern = r'<br><br><strong>Specifications:</strong>'
        end_pattern = r'<br><br><strong>Package Includes:</strong><br>'
    
        # Find starting and ending positions using regular expressions
        start_pos = re.search(re.escape(start_pattern), text)
        end_pos = re.search(re.escape(end_pattern), text)
    
        # If both start and end patterns are found
        if start_pos and end_pos:
            # Extract the substring between start and end patterns
            substring_between_patterns = text[start_pos.end():end_pos.start()]
    
            # Count occurrences of '<br>•'
            count_br_bullet = substring_between_patterns.count('<br>•')
    
            # If the count is less than 2, replace the text between patterns with the end pattern
            if count_br_bullet < 2:
                text = text[:start_pos.start()] + end_pattern + text[end_pos.end():]
    
        return text


    # Apply the function to the 'Body HTML' column
    df3['Body HTML'] = df3['Body HTML'].apply(replace_text_between_patterns)

    
    
    
    ### HTML: 
    output_file_path = initial_raw_file.replace("RawExport", "FinalCleanData")
    
    df3.to_excel(output_file_path, index=False)
    
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
        for index, row in df3.iterrows():
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