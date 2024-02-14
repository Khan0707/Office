import pandas as pd
import regex as re
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")#, category=DeprecationWarning)
import time



python_script_cleaned = r"D:\BackendData\Idropship\12_12_Idropship_CleanedDesc_NZ_HTML.xlsx"

initial_raw_file= r"D:\BackendData\Idropship\12_12_Idropship_RawExport_NZ.xlsx"


# Read the Excel file into a DataFrame
df = pd.read_excel(python_script_cleaned)
df2= pd.read_excel(initial_raw_file)
df2 = df2[df2['Body HTML'].notna()]

print("*********************** Different Category Types in the Data :*******************************")
value_counts_with_index = df2.groupby('Type').apply(lambda x: pd.Series({'count': x['Type'].count(), 'start_index': x.index[0]}))
print(value_counts_with_index)
print("***********************    *******************************")

time.sleep(2) 

print("**********************Delayed message.*******************************")
# Extract initial word
df['Initial_Word'] = df['Body HTML'].str.split().str[0]
# Count occurrences
word_counts = df['Initial_Word'].value_counts()
print(f"Starting Words are : {word_counts}")
time.sleep(2) 
print("**********************Delayed message.*******************************")






def features_section_check(text):
    if pd.notna(text):
        target_string_lowercase = 'key features'
        regex_pattern = re.compile(re.escape(target_string_lowercase), re.IGNORECASE)

        if "<strong>features:</strong><br>" not in text.lower() and regex_pattern.search(text):
            text = regex_pattern.sub("<strong>Features:</strong><br>", text)
            return text
        else:
            return text



df['Body HTML'] = df['Body HTML'].apply(lambda x: features_section_check(x))


# Function to delete everything before "<strong>Features: </strong><br>" i.e removing deescriptions
def delete_before_string(text, target_string):
    if pd.notna(text):
        if "<strong>Features:</strong><br>" in text:
            index = text.find("<strong>Features:</strong><br>")
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

# Bullets:
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
                    features_to_specs = text[features_index:specs_index].replace('<br>  •', '<br>')
                    specs_to_package = text[specs_index:package_index].replace('<br>  •', '<br>')
                    package_to_end = text[package_index:].replace('<br>  •', '<br>')
                    # Replace <br> and <br> • with <br> • between the specified keywords
                   #text = text[features_index:specs_index].replace('<br> ', '<br>  •').replace('<br> •', '<br>  •') + text[specs_index:package_index].replace('<br>', '<br>  •').replace('<br> •', '<br>  •') + text[package_index:].replace('<br>', '<br>  •').replace('<br> •', '<br>  •')
                    text = (
                        features_to_specs.replace('<br>', '<br>•') +
                        specs_to_package.replace('<br>', '<br>•') +
                        package_to_end.replace('<br>', '<br>•')
                    )

    return text

# Function to replace <br> • • with <br> • in the entire DataFrame
def replace_double_br_space(df):
    df.replace('<br>• •', '<br>•', inplace=True, regex=True)
    return df


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
    

# Apply the function to the 'Body HTML' column
df['Body HTML'] = df['Body HTML'].apply(lambda x: delete_before_string(x, "<strong>Features:</strong><br>"))

df['Body HTML'] = df['Body HTML'].str.replace('<br>: ','<br>')
print("**********************Delayed message.*******************************")
# Extract initial word
df['Initial_Word'] = df['Body HTML'].str.split().str[0]
# Count occurrences
word_counts = df['Initial_Word'].value_counts()
print(f"Starting Words Post Deleting Description are : {word_counts}")
time.sleep(2) 
print("**********************Delayed message.*******************************")




df['Body HTML'] = df['Body HTML'].apply(replace_br_between_keywords)


# Apply the function to the entire DataFrame
df.apply(replace_double_br_space)
df['Body HTML'] = df['Body HTML'].apply(remove_br_before_strong_and_between_br)


indexes_without_package_includes = df[~df['Body HTML'].str.contains('Package Includes:', case=False)].index
print(f"RawData_without_package_includes: {len(indexes_without_package_includes)}")
print(f"RawData_without_package_includes: {indexes_without_package_includes}")


# Merge the DataFrames on 'Variant SKU' and replace 'Body HTML' in df2 with 'Body HTML' in df
df3 = pd.merge(df2, df[['Variant SKU', 'Body HTML']], on='Variant SKU', how='left', suffixes=('_replace', '_original'))

# Use 'Body HTML_original' as the final 'Body HTML' column in df3
df3['Body HTML'] = df3['Body HTML_original'].combine_first(df3['Body HTML_replace'])

# Drop the unnecessary columns
df3 = df3.drop(['Body HTML_replace', 'Body HTML_original'], axis=1)
# Set the column order to match df2
column_order = df2.columns
df3 = df3[column_order]


# Convert Inches to cm in Title and HTML: 
    
def convert_inches_to_cm(text):
    # Regular expression to find inches pattern
    inches_pattern = r'(\d+(\.\d+)?)\s*\"'

    def inches_to_cm(match):
        inches_value = float(match.group(1))
        cm_value = round(inches_value * 2.54)
        return f"{cm_value} cm"

    # Using re.sub with a function to replace inches with cm
    text = re.sub(inches_pattern, inches_to_cm, text)

    return text
df3 ['Body HTML'] = df3 ['Body HTML'].apply(convert_inches_to_cm)
df3 ['Title'] = df3 ['Title'].apply(convert_inches_to_cm)


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

# Check if 'not_update_CA' is present in Tags with the value in Vendor column

current_date = datetime.now().strftime('%d_%m')
df3_result['Tags']= df3_result['Tags'].apply(lambda x: x.replace('not_update_CA', f'{df3_result["Vendor"].iloc[0]}_{current_date}') if 'not_update_CA' in x else x)

df3_result['Tags'] =df3_result['Tags'].str.replace(r'color', 'Colour', case=False)

# Replace values in the Vendor column and store in a dictionary
vendor_mapping = {'idropship': 'GODIAU'}
df3_result['Vendor'] = df3_result['Vendor'].replace(vendor_mapping)

# Step 3: Replace values in other columns
df3_result['Status'] = df3_result['Status'].replace({'Draft': 'Active'})
df3_result['Published'] = df3_result['Published'].replace({False: True})
df3_result['Published Scope'] = df3_result['Published Scope'].replace({'web': 'global'})

# Convert a set of columns from numeric to text format
columns_to_convert = ['ID', 'Variant Inventory Item ID', 'Variant ID','Variant Barcode']

df3_result[columns_to_convert] =df3_result[columns_to_convert].astype(str)

#df3_result.drop('Variant Barcode', axis=1, inplace=True)
#print("Dropped Barcode Column")

# Capitalise lines of all 3 section:
def capitalize_sentences(text, custom_stop_words=None):
    if custom_stop_words is None:
        custom_stop_words = set(['and', '-and', 'the', 'an', 'of', 'is', 'in', 'to', 'for', 'with', 'X', 'With','on', 'from', 'with', 'a', 'as', 'kg ', 'cm ', 'x', 'are', 'so', 'that', ' m '])

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


df3_result['Body HTML'] = df3_result['Body HTML'].apply(capitalize_sentences)
df3_result['Body HTML']= df3_result['Body HTML'].str.replace("1 X","1 x")
df3_result['Body HTML']= df3_result['Body HTML'].str.replace("1X","1 x")
df3_result['Body HTML'] = df3_result['Body HTML'].replace({"\s*X\s*": " x ", "\s*Cm\s*": " cm ",  "\s*Mm\s*": " mm ","\s*Kg\s*": " kg "}, regex=True)

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



df3_result['Body HTML']= df3_result['Body HTML'].str.replace('<Br>• Complies With As 1892. 1: 2018 And En131 Safety Standards','')


df3_result['Body HTML']= df3_result['Body HTML'].str.replace('(Non-Foldable)<Br>•', '(Non-Foldable) ')

#### HTML: 
output_file_path = initial_raw_file.replace("RawExport", "FinalCleanData")

df3_result.to_excel(output_file_path, index=False)

def check_character_count(html):
    character_count = len(html)
    status = 'Limit Crossed' if character_count > 1600 else 'Limit Not Crossed'
    status_style = 'color: red; font-weight: bold; font-style: italic;' if status == 'Limit Crossed' else ''
    character_count_style = 'color: red; font-weight: bold; font-style: italic;' if character_count > 1600 else ''
    return status, status_style, character_count_style
        
finalhtml = output_file_path.replace('xlsx', 'html')

print("Creating HTML File ....")    
   
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