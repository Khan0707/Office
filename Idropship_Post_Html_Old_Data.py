import pandas as pd
import regex as re
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")#, category=DeprecationWarning)
import time



python_script_cleaned = r"D:\BackendData\Idropship\13_12_Idropship_Old_CleanedDesc_NZ_HTML.xlsx"

initial_raw_file= r"D:\BackendData\Idropship\13_12_Idropship_Old_RawExport_NZ.xlsx"

# Read the Excel file into a DataFrame
df = pd.read_excel(python_script_cleaned, engine='openpyxl')
df2= pd.read_excel(initial_raw_file)
df2 = df2[df2['Body HTML'].notna()]

pattern = re.compile(r'[â€¢]|_x000D_|ï¼Š|ï¼š|: •')#
df['Body HTML'] = df['Body HTML'].apply(lambda x: pattern.sub('', str(x)))

# Define a regular expression pattern to match consecutive occurrences of •
pattern = re.compile(r'•+')
df['Body HTML'] = df['Body HTML'].apply(lambda x: pattern.sub('•', str(x)))

df['Body HTML'] = df['Body HTML'].str.replace('<br> SPECIFICIATION<br>' , '<br><strong>Specifications:</strong><br><br>')
df['Body HTML'] = df['Body HTML'].str.replace('<br> <br>', '<br>')

time.sleep(2) 

print("**********************Delayed message.*******************************")
# Extract initial word
df['Initial_Word'] = df['Body HTML'].str.split().str[0]
# Count occurrences
word_counts = df['Initial_Word'].value_counts()
print(f"Starting Words are : {word_counts}")
time.sleep(2) 
print("**********************Delayed message.*******************************")


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
   , 'delivery includes':'<br><br><strong>Package Includes:</strong>'
   , 'delivering includes' :'<br><br><strong>Package Includes:</strong>'
   , 'The tool set includes:': '<br><br><strong>Package Includes:</strong>'
    
}

for pattern, replacement in package_pattern.items():
    df['Body HTML'] = df['Body HTML'].str.replace(pattern, replacement, regex=True, flags=re.IGNORECASE)

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
        elif "Description<br>" in text:
            return text.replace("Description<br>", "<strong>Features:</strong><br>")
        elif "><strong>Package Includes:</strong>" in text:
            index = text.find("><strong>Package Includes:</strong>")
            if index != -1:
                return text[index:]
    return text


# Apply the function to the 'Body HTML' column
df['Body HTML'] = df['Body HTML'].apply(lambda x: delete_before_string(x, "<strong>Features:</strong><br>"))

df['Body HTML'] = df['Body HTML'].str.replace('<br>: ','<br>')

additional_patterns = [
    '<br>ï¼š<br>',
    '<br>Ã¯Â¼Å¡<br>',
    '<br>?<br>', '<br>?<br>',
    '<br>T<br>'
]
# Replace additional patterns
for pattern in additional_patterns:
    df['Body HTML'] = df['Body HTML'].str.replace(pattern, '<br><br>', regex=True, flags=re.IGNORECASE)

remove_patterns = [
    r'\?Frame\+Desktop\?',
    r'Noteï¼Š4 Types For Choice, Sold Seperately!',
    r'Note\?4 Types For Choice, Sold Seperately!',
    r'Note4 Types For Choice, Sold Seperately!',
    r'Key Feature S'
]

# Replace additional patterns
for pattern in remove_patterns:
    df['Body HTML'] = df['Body HTML'].str.replace(pattern, '', regex=True, flags=re.IGNORECASE)



print("**********************Delayed message.*******************************")
# Extract initial word
df['Initial_Word'] = df['Body HTML'].str.split().str[0]

df = df[df['Initial_Word'] != '><strong>Package'].reset_index(drop=True)

# Count occurrences
word_counts = df['Initial_Word'].value_counts()
print(f"Starting Words Post Deleting Description are : {word_counts}")
time.sleep(2) 
print("**********************Delayed message.*******************************")


   
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
    







df['Body HTML'] = df['Body HTML'].apply(replace_br_between_keywords)


# Apply the function to the entire DataFrame
df.apply(replace_double_br_space)
df['Body HTML'] = df['Body HTML'].apply(remove_br_before_strong_and_between_br)


indexes_without_package_includes = df[~df['Body HTML'].str.contains('Package Includes:', case=False)].index
print(f"RawData_without_package_includes: {len(indexes_without_package_includes)}")
print(f"RawData_without_package_includes: {indexes_without_package_includes}")


# Merge the DataFrames on 'Variant SKU' and replace 'Body HTML' in df2 with 'Body HTML' in df
# Merge the DataFrames on 'Variant SKU' and replace 'Body HTML' in df2 with 'Body HTML' in df
df3 = pd.merge(df2, df[['Variant SKU', 'Body HTML']], on='Variant SKU', how='inner', suffixes=('_replace', '_original'))

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
        brand_from_tags = row['Brand'].strip().lower() if pd.notna(row['Brand']) else 'NA'
        
        # Extract the brand from Title column
        brand_from_title = row['Title'].split()[0].lower()

        # Check if the brand from Tags matches the brand from Title
        if brand_from_tags == brand_from_title:
            # Drop the word following 'Brand_' from all cells of Title column
            # brand_length = len('Brand_')
            df.at[index, 'Title'] = row['Title'].lower().replace(f'{brand_from_title}', '', 1).strip()

        

    # Replicate the value of Title to Handle column in lowercase using '-'
    #df['Handle'] = df['Title'].str.lower().str.replace(' ', '-')
    df['Title'] = df['Title'].apply(lambda x: x.title())
    return df  # Return the modified DataFrame

# Call the function with df3 and capture the result
df3_result = process_df3(df3)


def remove_large_highlights(row):
    start_pattern = r'<strong>Features:</strong>'
    end_pattern = r'<strong>Specifications:</strong>'
    
    # Extract text between start and end patterns
    match = re.search(f'{start_pattern}(.*?){end_pattern}', row, re.DOTALL)
    if match:
        highlighted_text = match.group(1)
        
        # Check if there is only '<br><br>' and no other text, and if character count exceeds 1600
        if len(row) > 1500:
            # Remove the entire block
            return row.replace(f'{start_pattern}{highlighted_text}{end_pattern}','<strong>Specifications:</strong><br>')
    return row


df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_large_highlights)


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


def make_tags_lowercase(text):
    patterns = [r'<Br>', r'<Strong>', r'</Strong>']

    for pattern in patterns:
        text = re.sub(re.escape(pattern), pattern.lower(), text)

    return text


df3_result['Body HTML'] = df3_result['Body HTML'].apply(capitalize_sentences)

df3_result['Body HTML'] = df3_result['Body HTML'].apply(make_tags_lowercase)


df3_result['Body HTML']= df3_result['Body HTML'].str.replace("1 X","1 x")
df3_result['Body HTML']= df3_result['Body HTML'].str.replace("1X","1 x")

df3_result['Body HTML'] = df3_result['Body HTML'].replace(
    {"\\s*X\\s*": " x ", "\\s*Cm\\s*": " cm ", "\\s*Mm\\s*": " mm ", "KGS": " kgs", "\\s*Kg\\s*": " kg "},
    regex=True
)


# Create a new column 'Features_Line' with the brand information
df3_result['Features_Line'] = '<strong>Features:</strong><br>• Brand : ' + df3_result['Brand'].astype(str)

# Function to replace '<strong>Features:</strong>' line in 'Body HTML'
def replace_features_line(row):
    if pd.notnull(row['Body HTML']):
        return re.sub(r'<strong>Features:</strong>', row['Features_Line'], row['Body HTML'])
    return row['Body HTML']

# Apply the replacement function to each row
df3_result['Body HTML'] = df3_result.apply(replace_features_line, axis=1)
df3_result['Body HTML']= df3_result['Body HTML'].str.replace('<br>• Complies With As 1892. 1: 2018 And En131 Safety Standards','')


df3_result['Body HTML']= df3_result['Body HTML'].str.replace('(Non-Foldable)<br>•', '(Non-Foldable) ')




def replace_text_between_keywords_sizes(df):
    bed_sizes = ["Single", "King Single", "Double", "Queen", "King", "Super King"]
    
    for size in bed_sizes:
        mask = df['Title'].str.contains(size, case=False)
        pattern = r'<br>• Dimensions.*?(?=<br><br><strong>Specifications:</strong>)'
        
        df.loc[mask, 'Body HTML'] = df.loc[mask, 'Body HTML'].str.replace(pattern, '', regex=True)

# Example usage:
replace_text_between_keywords_sizes(df3_result)




count_multiple_occurrences = (df3_result['Body HTML'].str.count('Package Includes:') > 1).sum()
print(f"Count of rows with more than one occurrence of 'Package Includes:' in 'Body HTML': {count_multiple_occurrences}")

indices_multiple_occurrences = df.index[df['Body HTML'].str.count('Package Includes:') > 1].tolist()
print(f"Indices of rows with more than one occurrence of 'Package Includes:' in 'Body HTML': {indices_multiple_occurrences}")

def replace_text_between_keywords(df, keyword1, keyword2):
    df['Body HTML'] = df['Body HTML'].replace(f'{keyword1}.*?{keyword2}', f'{keyword2}', regex=True)


replace_text_between_keywords(df3_result, '<br><strong>Package Includes:</strong>', '<br><strong>Package Includes:</strong>')

replace_text_between_keywords(df3_result, '<br>• Tile Leveling System Clips:','<br>• Tile Sucker Set:')


replace_text_between_keywords(df3_result, '<br>• IMPORTANT','<br><br><strong>Package Includes:</strong>')

def convert_pattern_to_space(text):
    # Use regular expression to find patterns like '1X', '1 x', etc.
    pattern = re.compile(r'(\d+)\s*[xX]\s*', re.IGNORECASE)

    # Replace the found patterns with '1 x', '2 x', etc.
    result = pattern.sub(r'\1 x ', text)
    
    return result
df3_result['Body HTML'] = df3_result['Body HTML'].apply(convert_pattern_to_space)

# Convert commas in PAckage Includes into Bullets
df3_result['Body HTML']= df3_result['Body HTML'].apply(lambda x: re.sub(r'(?i)(?<=<strong>Package Includes:</strong>.*?)(,\s*|\s*<br>\s*•\s*)', '<br>• ', x.title()))

# Convert + sign in PAckage Includes into Bullets
df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: re.sub(r'(?i)(?<=<strong>Package Includes:</strong>.*?)\+', '<br>• ', x.title()))

df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: re.sub(r'(?i)(?<=<strong>Package Includes:</strong>.*?)Pcs', 'Pcs:', x.title()))

df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: re.sub(r'(?i)(?<=<strong>Package Includes:</strong>.*?):\s*', ': <br>• ', x.title()))




df3_result['Body HTML'] = df3_result['Body HTML'].apply(make_tags_lowercase)

df3_result['Body HTML'] = df3_result['Body HTML'].replace({"\s*X\s*": " x ", "\s*Cm\s*": " cm ",  "\s*Mm\s*": " mm ","\s*Kg\s*": " kg "}, regex=True)

df3_result['Title'] = df3_result['Title'].replace({"\s*X\s*": " x ", "\s*Cm\s*": " cm ",  "\s*Mm\s*": " mm ","\s*Kg\s*": " kg "}, regex=True)

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(' kg S', ' kg ')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('x l','XL')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('x  x xl','XXXL', case=False)
df3_result['Title'] = df3_result['Title'].str.replace('x l','XL')
df3_result['Title'] = df3_result['Title'].str.replace('x xl','XXL')
df3_result['Title'] = df3_result['Title'].str.replace('x xxl','XXXL')

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('x xl','XXL')

indexes_without_package_includes = df3_result[~df3_result['Body HTML'].str.contains('Package Includes:', case=False)].index
print(f"FinalData_without_package_includes: {len(indexes_without_package_includes)}")
print(f"FinalData_without_package_includes: {indexes_without_package_includes}")

# Space between Decimals
df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: re.sub(r'(\d+)\s*\.\s*(\d+)', r'\1.\2', x))

# Manual Word Replaement :
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('Aa Batteries ', 'AA Batteries') 

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(r'x /' , r'x ')


remove_patterns2 = [
    r'x L: ' , r'• 66M', r'Ãžæ’', r'x  XL:', r'x  x xl:', '<br>• Brand : Nan',r'Ï¼®', r'<br>• Key Feature', r'<br>• Key F Eatures'
    , '<br>• Key Fea Tur Es', r'ï¼', r'<br>• Brandlevede', r'<br>• Brand\? Levede', r'<br>• Branddreamz', r'<br>• :' , r'Ãž¡', r'<br>• Easy To Disposal Of'
]

# Replace additional patterns
for pattern in remove_patterns2:
    df3_result['Body HTML']  = df3_result['Body HTML'] .str.replace(pattern, '', regex=True, flags=re.IGNORECASE)

def replace_text_between_keywords_element(df, index, keyword1, keyword2):
    df['Body HTML'].iloc[index] = re.sub(f'{keyword1}.*?{keyword2}', f'{keyword2}', df['Body HTML'].iloc[index], flags=re.DOTALL)

replace_text_between_keywords_element(df3_result,2669, '• Small','• Large :')

replace_text_between_keywords_element(df3_result,34, '• Medium :','<br><br><strong>Package Includes:')

replace_text_between_keywords_element(df3_result,319, '• Large :','<br><br><strong>Package Includes:')
replace_text_between_keywords_element(df3_result,319, '• Small','• Medium :')

df3_result['Body HTML']  = df3_result['Body HTML'] .str.replace('<br>• <br>•','<br>•')
df3_result['Body HTML']  = df3_result['Body HTML'] .str.replace('<br> <br> •', '<br> •')

df3_result['Body HTML']  = df3_result['Body HTML'] .str.replace('kg s','kg')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(r'\'S ', 's', regex=True)

# Define the patterns
patterns = ['• Small :', '• Medium :', '• Large :']

# Create a regex pattern for matching the text between Specifications and Package Includes
specifications_pattern = re.compile(r'<strong>Specifications:</strong>(.*?)(?=<strong>Package Includes:</strong>)', re.DOTALL | re.IGNORECASE)

# Function to check if all patterns exist in the given text
def check_patterns(text):
    return all(re.search(pattern, text) for pattern in patterns)


matching_rows = df3_result[df3_result['Body HTML'].apply(lambda x: bool(re.search(specifications_pattern, x)))]

matching_indices = matching_rows[matching_rows.apply(lambda row: check_patterns(row['Body HTML']), axis=1)].index

print(f'Number of rows where all patterns exist: {len(matching_indices)}')
print(f'Indices of rows where all patterns exist: {matching_indices}')



df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(r'<br><br><strong>Specifications:</strong><br><br><strong>Package Includes:</strong>','<br><br><strong>Package Includes:</strong>')

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(': <br>• x Pet Training Pads', ' x Pet Training Pads')



#### HTML: 
df3_result.drop([29,31,58,934,1134,2746,2344, 4020, 4065], inplace=True)
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