# -*- coding: utf-8 -*-
"""
Created on Tue Dec 19 07:51:38 2023

@author: aaqib
"""

import pandas as pd
import regex as re
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")#, category=DeprecationWarning)
import time

python_script_cleaned_base = r"D:\BackendData\Fragrancex\19_12_Fragrancex_CleanedDesc_xyz_HTML.xlsx"
initial_raw_file_base= r"D:\BackendData\Fragrancex\19_12_Fragrancex_RawExport_xyz.xlsx"

# # List of countries
countries = ['AU','NZ','GKZ'] 
#country='NZ'
for country in countries:
    python_script_cleaned = python_script_cleaned_base.replace('xyz', country)
    initial_raw_file = initial_raw_file_base.replace('xyz', country)
    
    # Read the Excel file into a DataFrame
    df = pd.read_excel(python_script_cleaned)
    df2= pd.read_excel(initial_raw_file)
    df2 = df2[df2['Body HTML'].notna()]
    
    # Extract initial word
    df['Initial_Word'] = df['Body HTML'].str.split().str[0]
    
    # Count occurrences
    word_counts = df['Initial_Word'].value_counts()
    
    # Sort by 'Initial_Word'
    df_sorted = df.sort_values(by='Initial_Word')
    
    df['Body HTML'] = '<strong>Highlights:</strong><br> ' + df['Body HTML']
    
    indexes_without_package_includes = df[~df['Body HTML'].str.contains('Package Includes:', case=False)].index
    print(f"Count_indexes_without_package_includes: {len(indexes_without_package_includes)}")
    print(f"indexes_without_package_includes: {indexes_without_package_includes}")
    
    
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
            brand_from_tags = row['Tags'].split(',')[0].strip().lower().replace('brand_','').split(' ')[0]
    
            # Extract the brand from Title column
            brand_from_title = row['Title'].split()[0].lower()
    
            # Check if the brand from Tags matches the brand from Title
            if brand_from_tags == brand_from_title:
                # Drop the word following 'Brand_' from all cells of Title column
                # brand_length = len('Brand_')
                df.at[index, 'Title'] = row['Title'].lower().replace(f'{brand_from_title}', '', 1).strip()
    
    
        # Replicate the value of Title to Handle column in lowercase using '-'
        df['Handle'] = df['Title'].str.lower().str.replace(' ', '-')
        df['Title'] = df['Title'].apply(lambda x: x.title())
        return df  # Return the modified DataFrame
    
    # Call the function with df3 and capture the result
    df3_result = process_df3(df3)
    
    
    current_date = datetime.now().strftime('%d_%m')
    
    #df3_result['Tags'] = df3_result['Tags'].apply(lambda x: x.replace('not_update_CA', f'{df3_result["Vendor"].iloc[0]}_{current_date}') if 'not_update_CA' in x else x)
    
    # Check if 'Variant Image' is not null before applying the lambda function
    df3_result['Tags'] = df3_result.apply(lambda row: row['Tags'].replace('not_update_CA', f'{df3_result["Vendor"].iloc[0]}_{current_date}') 
                                          if (not pd.isnull(row['Variant Image'])) and ('not_update_CA' in row['Tags']) 
                                          else row['Tags'], axis=1)
    
    df3_result['Tags'] = df3_result['Tags'].str.replace(r'color', 'Colour', case=False)
    
    
    if country == 'NZ':
        df3_result['Vendor'] = 'GOUS'
    
    elif country == 'AU':
        df3_result['Vendor'] = 'PDUS'
        
    elif country == 'GKZ':
    
        tags_vendor = df3_result['Tags'].str.extract(r'Brand_(.*?),')
        df3_result['Vendor'] = tags_vendor if not tags_vendor.empty else 'GoKinzo'
        df3_result['Tags'] = df3_result['Tags']+ 'FXUS, Location_USA'
        
    
    
    # Step 3: Replace values in other columns
    df3_result['Status'] = df3_result['Status'].replace({'Draft': 'Active'})
    df3_result['Published'] = df3_result['Published'].replace({False: True})
    df3_result['Published Scope'] = df3_result['Published Scope'].replace({'web': 'global'})
    
    # Convert a set of columns from numeric to text format
    columns_to_convert = ['ID', 'Variant Inventory Item ID', 'Variant ID', 'Variant Barcode']
    df3_result[columns_to_convert] =df3_result[columns_to_convert].astype(str)
    
    
    ## Adding BrandName to Specs:
        
    def extract_brand(tags):
        if pd.notnull(tags):
            brand_match = re.search(r'Brand_(.*?),', tags)
            if brand_match:
                return brand_match.group(1).split(',')[0].title()
        return None
    
    df3_result['Brand'] = df3_result['Tags'].apply(extract_brand)
    
    
    df3_result['Features_Line'] = '<strong>Specifications:</strong><br>• Brand : ' + df3_result['Brand'].astype(str)
    
    # Function to replace '<Strong>Features:</Strong>' line in 'Body HTML'
    def replace_features_line(row):
        if pd.notnull(row['Body HTML']):
            return re.sub(r'<strong>Specifications:</strong>', row['Features_Line'], row['Body HTML'])
        return row['Body HTML']
    
    # Apply the replacement function to each row
    df3_result['Body HTML'] = df3_result.apply(replace_features_line, axis=1)
    
    # Drop intermediate columns if needed
    df3_result = df3_result.drop(columns=['Brand', 'Features_Line'])
    
    tags_gender = df3_result['Tags'].str.extract(r'Gender_(.*?),')
    print(f"*** Gender in Tags: *** \n {tags_gender.value_counts()}")
    
    
    # Capitalise lines of Highlights section:
    def capitalize_sentences(text):
        start_pattern = r'<strong>Highlights:</strong><br>'
        end_pattern = r'<br><br><strong>Specifications:</strong><br>'
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
            final_text = [sentence.strip().capitalize() if sentence.lower().strip() not in custom_stop_words else sentence.strip() for sentence in sentences]
            capitalized_text = '. '.join(final_text)
    
            # Remove extra full stops after punctuation signs
            capitalized_text = re.sub(r'(?<=[.!?])\s*\.', '', capitalized_text)
    
            # Replace the original highlighted text with the modified version
            text = text[:start_pos] + capitalized_text + text[end_pos:]
    
        return text
    
    df3_result['Body HTML'] = df3_result['Body HTML'].apply(capitalize_sentences)
    
    
    # Add Bullets to all three sections:
    def replace_br_between_keywords(text):
        if pd.notna(text):
            highlight_index = text.find("<strong>Highlights:</strong>")
            if highlight_index != -1:
                specs_index = text.find("<strong>Specifications:</strong>")
                if specs_index != -1:
                    # Find the starting index of "<strong>Package Includes:</strong>"
                    package_index = text.find("<strong>Package Includes:</strong>")
                    if package_index != -1:
                        # Remove all occurrences of • between the specified keywords
                        highlights_specs= text[highlight_index:specs_index]
                        specs_to_package = text[specs_index:package_index].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>')
                        package_to_end = text[package_index:].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>')
                        # Replace <br> and <br> • with <br> • between the specified keywords
                       #text = text[features_index:specs_index].replace('<br> ', '<br>  •').replace('<br> •', '<br>  •') + text[specs_index:package_index].replace('<br>', '<br>  •').replace('<br> •', '<br>  •') + text[package_index:].replace('<br>', '<br>  •').replace('<br> •', '<br>  •')
                        text = (highlights_specs+#.replace('.', '.<br>') +
                            specs_to_package.replace('<br>', '<br>•') +
                            package_to_end.replace('<br>', '<br>•')
                        )
    
        return text
    
    df3_result['Body HTML'] = df3_result['Body HTML'].apply(replace_br_between_keywords)
    
    
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
    
    
    # Function to replace text when length is <= 100
    def replace_text_short_length(text):
        start_pattern = r'<strong>Highlights:</strong><br>'
        end_pattern = r'<br><br><strong>Specifications:</strong><br>'
    
        # Find the start and end positions
        start_pos = re.search(start_pattern, text)
        end_pos = re.search(end_pattern, text)
    
        if start_pos and end_pos:
            start_pos = start_pos.end()
            end_pos = end_pos.start()
    
            # Calculate the length of characters between start and end patterns
            length_between = end_pos - start_pos
    
            # Replace text if length is <= 100
            if length_between <= 100:
                text = text[end_pos:]
                text = text.replace(end_pattern,'<strong>Specifications:</strong><br>')
    
        return text
    
    # Apply the function to the 'Body HTML' column
    df3_result['Body HTML'] = df3_result['Body HTML'].apply(replace_text_short_length)
    
    
       
    



    
    def rearrange_content(text):
        # Define the pattern to extract content between Specifications and Package Includes
        pattern = re.compile(r'<strong>Specifications:</strong><br>(.*?)<br><br><strong>Package Includes:</strong>', re.DOTALL)
    
        # Find the match in the text
        match = pattern.search(text)
    
        # If there's a match, rearrange the content
        if match:
            specifications_content = match.group(1).strip()
    
            # Split the content into lines
            lines = specifications_content.split('<br>')
    
            # Sort the lines
            sorted_lines = reversed(lines)
    
            # Join the sorted lines back together
            rearranged_content = '<br>'.join(sorted_lines)
    
            # Replace the original content with the rearranged content
            updated_text = text.replace(specifications_content, rearranged_content)
    
            return updated_text
        else:
            return text

    # Check if the country is "GKZ"
    if country == 'GKZ':
        df3_result['Body HTML'] = df3_result.apply(lambda row: rearrange_content(row['Body HTML']), axis=1)
    
    
    
    
    
    
    
    
    
    
    
    ### Final Output:
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
        file.write('    max-width: 40%;\n')  # Decreased image size to 60%
        file.write('    height: auto;\n')
        file.write('}\n')
        file.write('p {\n')
        file.write('    line-height: 1.25;\n')  # Set line spacing to 1.0
        file.write('}\n')
        file.write('.body-html {\n')
        file.write('    font-size: normal;\n')
        file.write('    text-align: justify;\n')
        file.write('}\n')
        file.write('.missing-image {\n')
        file.write('    color: red;\n')
        file.write('    font-weight: bold;\n')
        file.write('}\n')
        file.write('</style>\n')
        file.write('</head>\n<body>\n')
    
        # Iterate through rows of the DataFrame
        for index, row in df3_result.iterrows():
            variant_sku = row['Variant SKU']
            title = row['Title']
            body_html = row['Body HTML']
            character_count_status, status_style, character_count_style = check_character_count(body_html)
    
            # Check if Variant Image URL is present
            variant_image_url = row['Variant Image']
            if pd.notnull(variant_image_url):
                # Write the image tag
                file.write(f'<main>\n')
                file.write(f'    <h3 style="font-size: larger;">S.No: {index} | Variant SKU: {variant_sku}</h3>\n')
                file.write(f'    <h2>Title: {title}</h2>\n')
                file.write(f'    <p style="font-size: normal; {character_count_style}">Character Count: <span style="{character_count_style}">{len(body_html)}</span> | Character Count Status: <span style="{status_style}">{character_count_status}</span></p>\n')
                file.write(f'    <img src="{variant_image_url}" alt="Variant Image">\n')
                file.write(f'    <div class="body-html">{body_html}</div>\n')
                file.write(f'</main>\n\n')
            else:
                # Highlight in red bold text if Variant Image URL is missing
                file.write(f'<main class="missing-image">\n')
                file.write(f'    <h3 style="font-size: larger; color: red; font-weight: bold;">S.No: {index} | Variant SKU: {variant_sku}</h3>\n')
                file.write(f'    <h2 style="color: red; font-weight: bold;">Title: {title}</h2>\n')
                file.write(f'    <p style="color: red; font-weight: bold;">Variant Image URL Missing</p>\n')
                file.write(f'    <p style="font-size: normal; {character_count_style}">Character Count: <span style="{character_count_style}">{len(body_html)}</span> | Character Count Status: <span style="{status_style}">{character_count_status}</span></p>\n')
                file.write(f'    <div class="body-html">{body_html}</div>\n')
                file.write(f'</main>\n\n')
    
        file.write('</body>\n</html>')
    
    print(f"HTML file :{finalhtml} created successfully.")
