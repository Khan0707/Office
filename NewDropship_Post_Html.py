# -*- coding: utf-8 -*-
"""
Created on Mon Nov 13 12:39:09 2023

@author: aaqib
"""

import pandas as pd
import regex as re
from datetime import datetime


python_script_cleaned = r"D:\BackendData\DropShipzone\21_12_Dropship_CleanedDesc_NZ_HTML.xlsx"

initial_raw_file= r"D:\BackendData\Dropshipzone\21_12_Dropship_RawExport_NZ.xlsx"

# Read the Excel file into a DataFrame
df = pd.read_excel(python_script_cleaned)
df2= pd.read_excel(initial_raw_file)
df2 = df2[df2['Body HTML'].notna()]

# Delete Description:
    
def delete_before_string(text, target_string):
    if pd.notna(text):
        if "<strong>Features:" in text:
            index = text.find("<strong>Features:")
            if index != -1:
                return text[index:]
        elif "<strong>Specifications:</strong>" in text:
            index = text.find("<strong>Specifications:</strong>")
            if index != -1:
                return text[index:]
        elif "><strong>Package Includes:</strong>" in text:
            index = text.find("><strong>Package Includes:</strong>")
            if index != -1:
                return text[index:] 
    return text


# remove data after a keyword section:
    
def remove_additional_details(input_string):
    split_result = input_string.split('<br> • Additional Details:', 1)
    
    if len(split_result) > 0:
        return split_result[0]
    else:
        return input_string
    

    
# Apply the function to the 'Body HTML' column
df['Body HTML'] = df['Body HTML'].apply(lambda x: delete_before_string(x, "<strong> Features:"))

df['Body HTML'] = df['Body HTML'].apply(remove_additional_details)

df['Body HTML'] = df['Body HTML'].replace('<br> Details', '<br><br><strong>Specifications:</strong>', regex=True)


# Merge the DataFrames on 'Variant SKU' and replace 'Body HTML' in df2 with 'Body HTML' in df
df3 = pd.merge(df2, df[['Variant SKU', 'Body HTML']], on='Variant SKU', how='left', suffixes=('_replace', '_original'))

# Use 'Body HTML_original' as the final 'Body HTML' column in df3
df3['Body HTML'] = df3['Body HTML_original'].combine_first(df3['Body HTML_replace'])

# Drop the unnecessary columns
df3 = df3.drop(['Body HTML_replace', 'Body HTML_original'], axis=1)
# Set the column order to match df2
column_order = df2.columns
df3 = df3[column_order]


# Manuplaing Handles  and TItle column, and emoving Brand NAme from Title
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

        # else:
        #     print(f"Title, Handle, and Tags are in order for SKU {row['Variant SKU']}.")

    # Replicate the value of Title to Handle column in lowercase using '-'
    df['Handle'] = df['Title'].str.lower().str.replace(' ', '-')
    df['Title'] = df['Title'].apply(lambda x: x.title())
    return df  # Return the modified DataFrame

# Call the function with df3 and capture the result
df3_result = process_df3(df3)

# Check if 'not_update_CA' is present in Tags with the value in Vendor column
current_date = datetime.now().strftime('%d_%m')
df3_result['Tags']= df3_result['Tags'].apply(lambda x: x.replace('not_update_CA', f'{df3_result["Vendor"].iloc[0]}_{current_date}') if 'not_update_CA' in x else x)

# Replace Color with Colour:
df3_result['Tags'] =df3_result['Tags'].str.replace(r'color', 'Colour', case=False)

# Replace values in other columns
df3_result['Status'] = df3_result['Status'].replace({'Draft': 'Active'})
df3_result['Published'] = df3_result['Published'].replace({False: True})
df3_result['Published Scope'] = df3_result['Published Scope'].replace({'web': 'global'})


# Convert a set of columns from numeric to text format
columns_to_convert = ['ID', 'Variant Inventory Item ID', 'Variant ID', 'Variant Barcode']
df3_result[columns_to_convert] =df3_result[columns_to_convert].astype(str)


# Replace values in the Vendor column based on the SKU_Variant column
df3_result['Vendor'] = df3_result.apply(lambda row: 'DZAUV' if row['Variant SKU'].startswith('V') else 'DZAU', axis=1)

# Replace </strong> with </strong><li> ; <strong> with </li><strong> ; then <br> <li> with 

## Adding Package Includes for those were it is missing:
def process_body_html(row):
    package_includes_str = '<br><strong>Package Includes:</strong><br> • 1 x '
    
    if '<br><br><strong>Package Includes:</strong><br>' not in row['Body HTML']:
        first_three_words = ' '.join(row['Title'].split()[:3]).upper()
        row['Body HTML'] += package_includes_str + first_three_words

    return row['Body HTML']

df3_result['Body HTML'] = df3_result.apply(process_body_html, axis=1) 

#removing numeric bullets:
def remove_number_bullets(text):
    # Define the pattern using regular expression
    pattern = r'\s*\(\d+\)\s*'

    # Use re.sub to remove the pattern from the text
    result = re.sub(pattern, '• ', text)

    return result
df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_number_bullets)

# Add Bullets to all three sections:
def replace_br_between_keywords(text):
    if pd.notna(text):
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
                    features_to_specs = text[features_index:specs_index].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>')
                    specs_to_package = text[specs_index:package_index].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>')
                    package_to_end = text[package_index:].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>')
                    # Replace <br> and <br> • with <br> • between the specified keywords
                   #text = text[features_index:specs_index].replace('<br> ', '<br>  •').replace('<br> •', '<br>  •') + text[specs_index:package_index].replace('<br>', '<br>  •').replace('<br> •', '<br>  •') + text[package_index:].replace('<br>', '<br>  •').replace('<br> •', '<br>  •')
                    text = (
                        features_to_specs.replace('<br>', '<br>•') +
                        specs_to_package.replace('<br>', '<br>•') +
                        package_to_end.replace('<br>', '<br>•')
                    )

    return text

df3_result['Body HTML'] = df3_result['Body HTML'].apply(replace_br_between_keywords)


def product_comes_with_patterns(text):
    if pd.notna(text):
        # Define the pattern for matching variations
        pattern = re.compile(r'''
            (?:This\s*item\s*comes\s*in\s*(\d+|\b(?:one|two|three|four|five|six|seven|eight|nine|ten)\b)\s*packages)|
            (?:Number\s*of\s*packages:\s*(\d+|\b(?:one|two|three|four|five|six|seven|eight|nine|ten)\b))|
            (?:No\.?\s*of\s*package[s]*:\s*(\d+|\b(?:one|two|three|four|five|six|seven|eight|nine|ten)\b))|
            (?:The\s*item\s*comes\s*in\s*(\d+|\b(?:one|two|three|four|five|six|seven|eight|nine|ten)\b)\s*package)
        ''', re.IGNORECASE | re.VERBOSE)

        # Search for the pattern in the text
        match = re.search(pattern, text)

        if match:
            # Fetch the number from the matched group
            number_spelling = next(group for group in match.groups() if group)

            # Create a mapping from words to digits
            word_to_digit = {
                'one': '1',
                'two': '2',
                'three': '3',
                'four': '4',
                'five': '5',
                'six': '6',
                'seven': '7',
                'eight': '8',
                'nine': '9',
                'ten': '10'
            }

            # Convert the spelling format to digit using the mapping
            number_digit = word_to_digit.get(number_spelling.lower(), number_spelling)

            # Replace the original pattern with the desired format
            replacement = 'This Product is Packaged and Delivers in {} Package'.format(number_digit)
            text = re.sub(pattern, replacement, text)

    return text

df3_result['Body HTML'] = df3_result['Body HTML'].apply(product_comes_with_patterns)


def syntax_package_includes(text):
    if pd.notna(text):
        # Find the starting index of "<strong>Package Includes:</strong>"
        package_index = text.find("<strong>Package Includes:</strong>")
        if package_index != -1:
            # Extract the text that comes after "<strong>Package Includes:</strong>"
            package_text = text[package_index + len("<strong>Package Includes:</strong>"):]

            # Apply the replacement only to the extracted package text
            package_text = re.sub(r'<br>• (.*?)(?:\s*x\s*(\d+|\b(?:one|two|three|four|five|six|seven|eight|nine|ten)\b))', r'<br>• \2 x \1', package_text, flags=re.IGNORECASE)

            # Replace the original package text with the modified version
            text = text[:package_index + len("<strong>Package Includes:</strong>")] + package_text

    return text


df3_result['Body HTML'] = df3_result['Body HTML'].apply(syntax_package_includes)



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




# df3_result['Body HTML'] = df3_result['Body HTML'].replace(['<strong>Package Includes:</strong><br>','<strong>Package Includes:</strong><br>','<strong>Package Includes:</strong>'], '<strong>Package Includes:</strong>', regex=True)

#df3_result['Body HTML'] = df3_result['Body HTML'].apply(replace_br_between_keywords)
#df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_br_before_strong_and_between_br)

patterns_to_replace = ['Please note', 'Important note', 'Pls Note']
replacement = 'Note'

for pattern in patterns_to_replace:
    df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(pattern, replacement, flags=re.IGNORECASE)
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('Note:', '', flags=re.IGNORECASE)
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('Use Manual', 'User Manual', flags=re.IGNORECASE)


#df3_result['Body HTML'].str.contains('<br><br><strong>Package Includes:</strong><br>').sum()


# Capitalise lines of Highlights section:
def capitalize_sentences(text,  start_pattern, end_pattern):
    
    # Define a custom list of stop words
    custom_stop_words = [
        'a', 'an', 'and', 'the', 'is', 'of', 'in', 'to', 'for', 'with', 'on', 'by', 'as',
        'at', 'but', 'or', 'not', 'from', 'so', 'that', 'it', 'this', 'these', 'those',
        'I', 'you', 'he', 'she', 'we', 'they', 'him', 'her', 'us', 'them', 'your', 'his',
        'its', 'our', 'their', 'which', 'who', 'whom', 'whose', 'what', 'when', 'where',
        'why', 'how', 'will', 'can', 'must', 'should', 'would', 'could', 'do', 'does',
        'did', 'doing', 'done', 'have', 'has', 'had', 'having', 'get', 'gets', 'got',
        'getting', 'been', 'be', 'am', 'isn\'t', 'aren\'t', 'wasn\'t', 'weren\'t', 'isn', 'aren',
        'wasn', 'weren', 'it\'s', 'that\'s', 'he\'s', 'she\'s', 'there\'s', 'here\'s',
        'i\'m', 'you\'re', 'they\'re', 'we\'re', 'i\'ll', 'you\'ll', 'he\'ll', 'she\'ll',
        'it\'ll', 'we\'ll', 'they\'ll', 'i\'d', 'you\'d', 'he\'d', 'she\'d', 'it\'d',
        'we\'d', 'they\'d', 'i\'ve', 'you\'ve', 'we\'ve', 'they\'ve', 'it\'s', 'that\'s',
        'let\'s', 'who\'s', 'what\'s', 'here', 'there', 'where', 'when', 'why', 'how'
    ]
    # Find the start and end positions
    start_pos = re.search(start_pattern, text)
    end_pos = re.search(end_pattern, text)

    if start_pos and end_pos:
        start_pos = start_pos.end()
        end_pos = end_pos.start()

        # Extract the text between the specified patterns
        highlighted_text = text[start_pos:end_pos]
        sentences = re.split(r'(?<=[.!?•])\s*', highlighted_text)
        
        final_text = [sentence.strip().lower() if sentence.lower().strip() in custom_stop_words 
              else sentence.strip().title() if (sentence.count(' ') <= 5  and sentence.lower().strip() not in custom_stop_words)  or ':' in sentence
              else sentence.strip().capitalize() 
              for sentence in sentences]
        


        final_text = ' '.join(final_text)

            
        
        # Remove extra full stops after punctuation signs
        final_text = re.sub(r'(?<=[.!?])\s*\.', '', final_text)

        # Replace the original highlighted text with the modified version
        text = text[:start_pos] + final_text + text[end_pos:]

    return text


start_pattern = r'<strong>Features:</strong><br>'
end_pattern = r'<br><br><strong>Specifications:</strong><br>'

df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: capitalize_sentences(x, start_pattern, end_pattern))




def make_tags_lowercase(text):
    patterns = [r'<Br>', r'<Strong>', r'</Strong>']

    for pattern in patterns:
        text = re.sub(re.escape(pattern), pattern.lower(), text)

    return text

df3_result['Body HTML'] = df3_result['Body HTML'].apply(make_tags_lowercase)

start_pattern = r'<br><br><strong>Specifications:</strong><br>'
end_pattern =  r'<br><br><strong>Package Includes:</strong><br>'
df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: capitalize_sentences(x, start_pattern, end_pattern))


df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(': Aa', ': AA', flags=re.IGNORECASE)

df3_result['Body HTML'] = df3_result['Body HTML'].apply(make_tags_lowercase)

def replace_and_remove(text, keyword):
    return re.sub(re.escape(keyword) + '.*$', '', text)

# Apply the function to the 'Body HTML' column
df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: replace_and_remove(x, "<br>• This product comes with"))


df3_result['Title'] = df3_result['Title'].replace({"Cm": "cm", "Pcs": "pcs"}, regex=True)
df3_result['Body HTML'] = df3_result['Body HTML'].replace({"Cm": "cm", "Pcs": "pcs", "Kg": "kg"}, regex=True)
df3_result['Body HTML'] = df3_result['Body HTML'].replace({"\s*X\s*": " x ", "\s*Cm\s*": " cm ", "\s*Mm\s*": " mm "}, regex=True)
df3_result['Body HTML'] = df3_result['Body HTML'].replace({"Men'S": "Men's", " x x box" : " x Xbox"}, regex=True)
df3_result['Body HTML'] = df3_result['Body HTML'].replace('i\.Pet', 'Pet', regex=True)
df3_result['Title'] = df3_result['Title'].replace('I\.Pet', 'Pet', regex=True)

# Space between Decimals
df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: re.sub(r'(\d+)\s*\.\s*(\d+)', r'\1.\2', x))

def remove_spaces(text):
    # Convert specific units to lowercase
    text = re.sub(r'\b(mm|cm|kg|inches)\b', lambda x: x.group().lower(), text, flags=re.IGNORECASE)
    
    # Remove spaces before ':', ',', 'mm', 'cm', 'kg', '('
    text = re.sub(r'\s+(:|,|mm|cm|kg|\()', r'\1', text, flags=re.IGNORECASE)
    
    # Remove spaces within expressions like 'w x h x d'
    #text = re.sub(r'\s*([wWhHdDxX])\s*([wWhHdDxX])\s*([wWhHdDxX])\s*', r'\1x\2x\3', text, flags=re.IGNORECASE)
    
    return text
df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_spaces)


indexes_without_package_includes = df3_result[~df3_result['Body HTML'].str.contains('Package Includes:', case=False)].index
print(f"Count_indexes_without_package_includes_at the END: {len(indexes_without_package_includes)}")
print(f"indexes_without_package_includes_at the END: {indexes_without_package_includes}")





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

