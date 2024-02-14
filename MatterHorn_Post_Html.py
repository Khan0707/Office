# -*- coding: utf-8 -*-
"""
Created on Wed Nov 22 07:30:55 2023

@author: aaqib
"""

import pandas as pd
import regex as re
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")#, category=DeprecationWarning)
import time

### NZ: 
    
# # List of countries
countries = ['AU','NZ']
dt= "11_12"

#country="NZ"
# Loop through each country

for country in countries:
#dd_mm format:


    python_script_cleaned_base = r"D:\BackendData\Matterhorn\abc_Matterhorn_CleanedDesc_xyz_HTML.xlsx"
    initial_raw_file_base= r"D:\BackendData\Matterhorn\abc_Matterhorn_RawExport_xyz.xlsx"
    
    print(" **** Country : ***")
    print(country)
    
    python_script_cleaned = python_script_cleaned_base.replace('xyz', country).replace('abc', dt)
    initial_raw_file = initial_raw_file_base.replace('xyz', country).replace('abc', dt)
    
    # Read the Excel file into a DataFrame
    df = pd.read_excel(python_script_cleaned)
    df2= pd.read_excel(initial_raw_file)
    df2 = df2[df2['Body HTML'].notna()]
    
    indexes_without_package_includes = df[~df['Body HTML'].str.contains('Specifications:', case=False)].index
    print(f"Count_indexes_without_package_includes: {len(indexes_without_package_includes)}")
    print(f"indexes_without_package_includes: {indexes_without_package_includes}")
    time.sleep(2) 
    print("**********************Delayed message.*******************************")
    # Extract initial word
    df['Initial_Word'] = df['Body HTML'].str.split().str[0]
    # Count occurrences
    word_counts = df['Initial_Word'].value_counts()
    print(f"Starting Words are : {word_counts}")
    time.sleep(2) 
    print("**********************Delayed message.*******************************")
    
    # Define a function to replace commas with periods in decimal numbers
    def replace_comma_with_period(text):
        pattern = r'(\d+,\d+)'
        return re.sub(pattern, lambda x: x.group().replace(',', '.'), text)

    
    # Apply the function to the 'Values' column
    df['Body HTML'] = df['Body HTML'].apply(replace_comma_with_period)
    
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
    
    
    # df['Body HTML'] = df['Body HTML'].str.replace() Dimensions: Width: 23 cm Height: 17 cm Depth: 9 cm
    replace_dict = {
    "<br> fabric<br>": "<br> fabric : ",
    "<br> warmers<br>": "<br> warmers :",
    'height<br>': 'height : ',
    "<br> Footbed<br>": "<br> Footbed :",
    # "Dimensions:" : "<br> Dimensions:",
    " Width:" : "<br> Width:",
    " Height:" : "<br> Height:",
    " Depth:" : "<br> Depth:",
    " Thickness:" : "<br> Thickness:",
    " Length:" : "<br> Length:",
    #r"\* Size <br>"  : 'Size',
    r":\s*\)" : ".",  
    r"\.{3,}": "."     # Replace three or more dots with a single dot
}

    df['Body HTML'] = df['Body HTML'].replace(replace_dict, regex=True)
   
     
    def sort_dimensions(text):
        start_pattern = r'<strong>Highlights:</strong><br>'
        end_pattern = r'<strong>Features:</strong><br>'
        spec_start = r'<br><strong>Specifications:</strong><br>'
        spec_end = r'<br><strong>Package Includes:</strong><br>'
        
        # Check if the text contains the specified patterns
        if start_pattern in text and end_pattern in text and "Dimensions" in text:
            highlights_match = re.search(f'{start_pattern}(.*?){end_pattern}', text, re.DOTALL)
            
            if highlights_match:
                # Extract the highlights text
                highlights_text = highlights_match.group(1).strip()
                
                # Extract the Dimensions text
                dimensions_match = re.search(r'Dimensions:(.*?)<br><br>', highlights_text, re.DOTALL)
                if dimensions_match:
                   
                    # dimensions_text = re.sub(r'At The Bottom', '', dimensions_match.group(1).strip(), flags=re.IGNORECASE)
                    # dimensions_text = re.sub(r'At The Top.', '', dimensions_text , flags=re.IGNORECASE)
                    # dimensions_text = re.sub(r'At The Widest Point', '', dimensions_text , flags=re.IGNORECASE)
                    #dimensions_text = dimensions_text.replace('cm ', ' cm <br> ')
                   
                    dimensions_text = dimensions_match.group(1).strip()#.replace('cm ', ' cm <br> ').replace('<br> at', 'at').replace('.', '. <br> ')#.replace(': <br>', '<br> ')
                    
                    # Replace the Dimensions text in the original highlights
                    updated_highlights_text = re.sub(r'Dimensions:(.*?)<br><br>', '', highlights_text, re.DOTALL)
                    
                    match = re.search(f'{re.escape(end_pattern)}(.*?){re.escape(spec_start)}', text, re.DOTALL)
                    if match:
                        extracted_text = match.group(1).strip()
        
                    # Append the updated highlights to Specifications
                    updated_specifications_text = f'{spec_start}{ dimensions_text}{spec_end}'
                    # Check if the text contains the specified pattern
                    if spec_end in text:
                        # Get the text after '<br><strong>Package Includes:</strong><br>'
                        text_after_package_includes_match = re.search(f'{re.escape(spec_end)}(.*?)$', text, re.DOTALL)
                        
                        if text_after_package_includes_match:
                            text_after_package_includes = text_after_package_includes_match.group(1).strip()
           
                        
                    # Replace the original highlights with updated specifications in the text
                    text = '<strong>Highlights:</strong><br><br>' + updated_highlights_text + \
                            '<br><br><strong>Features:</strong><br>' + extracted_text + \
                                updated_specifications_text + text_after_package_includes

        
        return text

    
    #df['Body HTML'] = df['Body HTML'].apply(sort_dimensions)


    def process_specifications(body_html):
        if '<br><br><strong>Specifications:</strong><br>' in body_html:
            extracted_text = body_html.split('<br><br><strong>Specifications:</strong><br>')[1]
            
            # Additional operations if the pattern exists
            extracted_text = re.sub(r'^\s*\d+\s*%\s*<br>\s*', '', extracted_text)
            extracted_text = re.sub(r'<br>\s*(\d+\s*%)', r': \1', extracted_text)
            
            # Replace patterns like "10.5 cm" with "10 cm"
            pattern = re.compile(r'(\d+\.\d+) cm')
            extracted_text = re.sub(pattern, lambda x: f'{float(x.group(1)):.0f} cm', extracted_text)
    
            # Custom function to add '*' before 'cm'
            def add_star(match):
                number_cm = match.group()
                if '*' not in number_cm:
                    return '* ' + number_cm
                else:
                    return number_cm
    
            # Regular expression to find the pattern 'number cm' without '*' before it
            pattern = re.compile(r'(?<=[^-/])\b\d+\s*cm\b')
    
            # Use re.sub with a custom replacement function
            extracted_text = re.sub(pattern, add_star, extracted_text)
    
            # Replace instances of '* *' with '*'
            extracted_text = re.sub(r'\*\s*\*', '*', extracted_text)
    
            # Format the text as specified
            lines = extracted_text.split('<br> *')
            formatted_lines = []
            for i, line in enumerate(lines):
                line = line.strip()
                if line != '':
                    if i > 1:
                        line = ' | ' + line
                    elif i <= 1:
                        line = '<br>• ' + line
                    formatted_lines.append(line)  
            
            output_text = ' '.join(formatted_lines)
            
            return output_text
        else:
            return 'NA'
    # Check if 'Specifications:' is present in 'Body HTML' and apply processing if true
    mask = df['Body HTML'].str.contains('Specifications:', case=False, na=False)
    
    df.loc[mask, 'processed_specifications'] = df.loc[mask, 'Body HTML'].apply(process_specifications)
    df.loc[mask, 'Body HTML'] = (
        df.loc[mask, 'Body HTML']
        .str.split('<br><br><strong>Specifications:</strong><br>')
        .str[0]
        + '<br><br><strong>Specifications:</strong><br>'
        + df.loc[mask, 'processed_specifications']
    )
    
    # df['processed_specifications'] = df['Body HTML'].apply(process_specifications)
    # df['Body HTML'] = df['Body HTML'].str.split('<br><br><strong>Specifications:</strong><br>').str[0] +'<br><br><strong>Specifications:</strong><br>'+ df['processed_specifications']
    
    # Merge the DataFrames on 'Variant SKU' and replace 'Body HTML' in df2 with 'Body HTML' in df
    df3 = pd.merge(df2, df[['Variant SKU', 'Body HTML']], on='Variant SKU', how='left', suffixes=('_replace', '_original'))
    # Use 'Body HTML_original' as the final 'Body HTML' column in df3
    df3['Body HTML'] = df3['Body HTML_original'].combine_first(df3['Body HTML_replace'])
    # Drop the unnecessary columns
    df3 = df3.drop(['Body HTML_replace', 'Body HTML_original'], axis=1)
    # Set the column order to match df2
    column_order = df2.columns
    df3 = df3[column_order]
    
    # TAgs and Titles:
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
    
    # Check if 'not_update_CA' is present in Tags with the value in Vendor column
    
    current_date = datetime.now().strftime('%d_%m')
    df3_result['Tags']= df3_result['Tags'].apply(lambda x: x.replace('not_update_CA', f'{df3_result["Vendor"].iloc[0]}_{current_date}') if 'not_update_CA' in x else x)
    
    df3_result['Tags'] =df3_result['Tags'].str.replace(r'color', 'Colour', case=False)
    
    # Apply the modifications directly to the DataFrame
    df3_result['Tags'] =df3_result['Tags'].apply(lambda tags: re.sub(r'Shipping_\d+\.\d+,', '', tags, flags=re.IGNORECASE))
    
    # Use regular expression to find the Type value in "Tags" and replace in the "Type" column
    df3_result['Type'] = df3_result['Tags'].apply(lambda tags: re.search(r'Type_(.+)', tags, flags=re.IGNORECASE).group(1) if re.search(r'Type_(.+)', tags, flags=re.IGNORECASE) else None)
    
    # Replace values in the Vendor column and store in a dictionary
    vendor_mapping = {'idropship': 'GODIAU', 'vidaxl':'GOAUAD', 'wefullfill':'Vibe Geeks', 'bigbuy':'PDBB', 'matterhorn':'GOEFASH'}
    
    if country in ['NZ', 'AU','US']:
        df3_result['Vendor'] = df3_result['Vendor'].replace(vendor_mapping)
    else: 
        #tags_vendor = df3_result['Tags'].str.split(',').str[0].str.strip().str.lower().str.replace('brand_','').str.split(' ').str[0].str.upper()
        tags_vendor = df3_result['Tags'].str.extract(r'Brand_(.*?),')
        df3_result['Vendor'] = tags_vendor if not tags_vendor.empty else 'GoSlash'
    # Step 3: Replace values in other columns
    df3_result['Status'] = df3_result['Status'].replace({'Draft': 'Active'})
    df3_result['Published'] = df3_result['Published'].replace({False: True})
    df3_result['Published Scope'] = df3_result['Published Scope'].replace({'web': 'global'})
    
    # Convert a set of columns from numeric to text format
    columns_to_convert = ['ID', 'Variant Inventory Item ID', 'Variant ID']
    df3_result[columns_to_convert] =df3_result[columns_to_convert].astype(str)
    
    df3_result.drop('Variant Barcode', axis=1, inplace=True)
    print("Dropped Barcode Column")
    
    
    value_counts_with_index = df3_result.groupby('Type').apply(lambda x: pd.Series({'count': x['Type'].count(), 'start_index': x.index[0]}))
    print(value_counts_with_index)
    indexes_without_package_includes = df3_result[~df3_result['Body HTML'].str.contains('Package Includes:', case=False)].index
    print(f"Count_indexes_without_package_includes: {len(indexes_without_package_includes)}")
    print(f"indexes_without_package_includes: {indexes_without_package_includes}")
    
    def remove_empty_highlights(row):
        start_pattern = r'<strong>Highlights:</strong><br>'
        end_pattern = r'<strong>Features:</strong><br>'
        
        # Extract text between start and end patterns (case-insensitive)
        match = re.search(f'{start_pattern}(.*?){end_pattern}', row, re.DOTALL | re.IGNORECASE)
        if match:
            highlighted_text = match.group(1)
            
            # Check if there is only '<br><br>' and no other text
            if highlighted_text.strip() == '':
                # Remove the entire block (case-insensitive)
                return re.sub(f'{start_pattern}.*?{end_pattern}', '<strong>Specifications:</strong><br>', row, flags=re.DOTALL | re.IGNORECASE)
        return row
    
    
    def remove_large_highlights(row):
        start_pattern = r'<strong>Highlights:</strong><br>'
        end_pattern = r'<strong>Features:</strong><br>'
        
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
    
    # Assuming df3_result is your DataFrame
    mask = df3_result['Type'] == 'Handbags'
    df3_result.loc[mask, 'Body HTML'] = df3_result.loc[mask, 'Body HTML'].apply(sort_dimensions)
    
    
    # Add Bullets to all three sections:
    def replace_br_between_keywords(text):
        if pd.notna(text):
            highlight_index = text.find("<strong>Highlights:</strong>")
            if highlight_index != -1:
            # Find the starting index of "<strong>Features: </strong>"
                features_index = text.find("<strong>Features:</strong>")
                if features_index != -1:
                    # Find the starting index of "<strong>Specifications:</strong>"
                    specs_index = text.find("<strong>Specifications:</strong>")
                    if specs_index != -1:
                        # Find the starting index of "<strong>Package Includes:</strong>"
                        package_index = text.find("<strong>Package Includes:</strong>")
                        if package_index != -1:
                            # Remove all occurrences of • between the specified keywords
                            highlights_specs= text[highlight_index:features_index]
                            features_to_specs = text[features_index:specs_index].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>')
                            specs_to_package = text[specs_index:package_index].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>')
                            package_to_end = text[package_index:].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>')
                            # Replace <br> and <br> • with <br> • between the specified keywords
                           #text = text[features_index:specs_index].replace('<br> ', '<br>  •').replace('<br> •', '<br>  •') + text[specs_index:package_index].replace('<br>', '<br>  •').replace('<br> •', '<br>  •') + text[package_index:].replace('<br>', '<br>  •').replace('<br> •', '<br>  •')
                            text = (highlights_specs+#.replace('.', '.<br>') +
                                features_to_specs.replace('<br>', '<br>•') +
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
    
    # Capitalise lines of all 3 section:
    def capitalize_sentences(text, custom_stop_words=None):
        if custom_stop_words is None:
            custom_stop_words = set(['and', '-and', 'the', 'an', 'of', 'is', 'in', 'to', 'for', 'with', 'X', 'With','on', 'from', 'with', 'a', 'as', 'kg', 'cm', 'x', 'are', 'so', 'that'])

        start_pattern = r'<strong>Features:</strong>'
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



    
    
    ## MANual Replacements:
    #df3_result['Body HTML'].str.contains("<br><br><br><br><br>").sum()
    df3_result['Body HTML']=df3_result['Body HTML'].str.replace("~","")
    df3_result['Body HTML']=df3_result['Body HTML'].str.replace("M<br>•","M |")
    df3_result['Body HTML']=df3_result['Body HTML'].str.replace("S<br>•","S |")
    df3_result['Body HTML']=df3_result['Body HTML'].str.replace("L<br>•","L |")
    df3_result['Body HTML']=df3_result['Body HTML'].str.replace("XL<br>•","XL |")
    df3_result['Body HTML']=df3_result['Body HTML'].str.replace("<br>• 78-86 cm","| 78-86 cm")
    
    ## : * 7 Cm shud be  : 7 Cm :
    df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(r': *\* ', ': ', regex= True)
    


    df3_result['Title']=df3_result['Title'].str.replace("By By","By")
    
    #CApitalise:
    df3_result['Body HTML'] = df3_result['Body HTML'].apply(capitalize_sentences)
    
    df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(r"`S", "'s", regex= True)
    df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(r"Cm","cm", regex= True)
    
    ## Adding Package Includes for those were it is missing:
    def add_package_includes(row):
        package_text = row['Title']
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
        
        elif '<br><br><strong>package includes:</strong>' not in row['Body HTML'].lower():
            
            text_title = re.split(r'(\d)', package_text, maxsplit=1)[0].strip()
            
            package_includes_str = '<br><strong>Package Includes:</strong><br> • 1 x '
            
            row['Body HTML'] += package_includes_str + text_title
            
            return row['Body HTML']
        else:
            return row['Body HTML'].replace('1 X ', '1 x ')
        
    df3_result['Body HTML'] = df3_result.apply(add_package_includes, axis=1)
    df3_result['Body HTML']=df3_result['Body HTML'].str.replace("<Br>• Cotton: 95 %: 5 % ", "<Br>• Cotton 95 % <Br>• Polyamide: 5 %", regex=True)
    df3_result['Body HTML']=df3_result['Body HTML'].str.replace("Xxl" , "XXL",regex=True)
    df3_result['Body HTML']=df3_result['Body HTML'].str.replace("Xl" , "XL",regex=True)
    df3_result['Body HTML']=df3_result['Body HTML'].str.replace("Xs" , "XS",regex=True)
    df3_result['Body HTML']=df3_result['Body HTML'].str.replace(r'<Br>• \* Size <Br>• ','<Br>• Size | ',regex=True)
    
    
   
    
    def extract_brand(tags):
        if pd.notnull(tags):
            brand_match = re.search(r'Brand_(.*?),', tags)
            if brand_match:
                return brand_match.group(1).split(',')[0].title()
        return None
    
    # Extract brand information from 'Tags' column
    df3_result['Brand'] = df3_result['Tags'].apply(extract_brand)
    
    # Create a new column 'Features_Line' with the brand information
    df3_result['Features_Line'] = '<strong>Features:</strong><br>• Brand : ' + df3_result['Brand'].astype(str)
    
    # Function to replace '<Strong>Features:</Strong>' line in 'Body HTML'
    def replace_features_line(row):
        if pd.notnull(row['Body HTML']):
            return re.sub(r'<strong>Features:</strong>', row['Features_Line'], row['Body HTML'])
        return row['Body HTML']
    
    # Apply the replacement function to each row
    df3_result['Body HTML'] = df3_result.apply(replace_features_line, axis=1)
    
    # Drop intermediate columns if needed
    df3_result = df3_result.drop(columns=['Brand', 'Features_Line'])


    
    
    
    
    
    
    
    
    
    
    
    
    
    
    ### HTML:
    
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