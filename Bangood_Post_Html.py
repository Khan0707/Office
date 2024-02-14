# -*- coding: utf-8 -*-
"""
Created on Thu Nov 23 21:27:19 2023

@author: aaqib
"""

import pandas as pd
import regex as re
from datetime import datetime
from text_cleaner import apply_removal_and_track_indices, words_to_remove



python_script_cleaned = r"D:\BackendData\Bangood\19_01_Bangood_CleanedDesc_NZ_HTML.xlsx"

initial_raw_file= r"D:\BackendData\Bangood\19_01_Bangood_RawExport_NZ.xlsx"

# Read the Excel file into a DataFrame
df = pd.read_excel(python_script_cleaned)
df2= pd.read_excel(initial_raw_file)
df2 = df2[df2['Body HTML'].notna()]

def extract_and_return_word_counts(dataframe):
    # Extract initial word
    dataframe['Initial_Word'] = dataframe['Body HTML'].str.split().str[0]

    # Count occurrences
    word_counts = dataframe['Initial_Word'].value_counts()

    # Sort DataFrame
    dataframe_sorted = dataframe.sort_values(by='Initial_Word', key=lambda x: x.map(word_counts), ascending=False)

    # Add Word_Counts column
    dataframe_sorted['Word_Counts'] = dataframe_sorted['Initial_Word'].map(word_counts)
    
    # Reset index
    dataframe_sorted = dataframe_sorted.reset_index(drop=True)


    return dataframe_sorted, word_counts

# Example usage:
    
df_sorted, counts = extract_and_return_word_counts(df)

df.loc[df['Initial_Word'] == '.size_table_start', 'Body HTML'] = df['Body HTML'].str.replace(r'^(.*?)\{clear:both \}<br> \*', r'<strong>Specifications:</strong><br>• ', regex=True)


df_sorted, counts = extract_and_return_word_counts(df)

df.loc[df['Initial_Word'] == "<br><br><strong>Specifications:</strong><br>", 'Body HTML'] = df['Body HTML'].str.replace(r"<br><br><strong>Specifications:</strong><br>", r'<strong>Specifications:</strong><br>• ', regex=True)

df.loc[df['Initial_Word'] == "<strong>Specification:</strong><br>", 'Body HTML'] = df['Body HTML'].str.replace(r"<strong>Specification:</strong><br>", r'<strong>Specifications:</strong><br>• ', regex=True)

df.loc[df['Initial_Word'] =='<br><br><strong>Features:</strong><br>', 'Body HTML'] = df['Body HTML'].str.replace(r'<br><br><strong>Features:</strong><br>', r'<strong>Features:</strong><br>', regex=True)


df_sorted, counts = extract_and_return_word_counts(df)

df['Body HTML'] = df['Body HTML'].apply(lambda x: x.split('<img loading=')[0] if '<img loading=' in x else x)
df['Body HTML'] = df['Body HTML'].apply(lambda x: x.split('<img alt=')[0] if '<img alt=' in x else x)

df_sorted, counts = extract_and_return_word_counts(df)


html_tags= ['<html>', '<head>', '<title>', '<body>', '<div>', '<p>', '<a>', '<img>', '<ul>', 
             '<li>', '<h1>', '<table>', '<tr>', '<td>', '<em>', '<span>','<hr>', 
             '<form>', '<input>', '<button>', '<select>', '<option>',
             '</html>', '</head>', '</title>', '</body>', '</div>', '</p>', '</a>',
             '</img>', '</ul>', '</li>', '</h1>', '</table>', '</tr>', '</td>']
# Create a regular expression pattern for all HTML tags
pattern = '|'.join(re.escape(tag) for tag in html_tags)

# Define a function to replace tags in a text
def replace_html_tags(text):
    return re.sub(pattern, '', text)

# Apply the function to the 'Body HTML' column
df['Body HTML'] = df['Body HTML'].apply(lambda x: replace_html_tags(x))




df = df[(df['Initial_Word'] == r'<strong>Specifications:</strong><br>•') | (df['Initial_Word'] == r'<strong>Features:</strong><br>')]


df_sorted, counts = extract_and_return_word_counts(df)


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
   '<br> Package included: <br>': '<br><br><strong>Package Includes:</strong><br>',
   'PRODUCT LIST': '<br><br><strong>Package Includes:</strong><br>',
   '<br> Package<br>': '<br><br><strong>Package Includes:</strong><br>',
   'Package included :': '<br><br><strong>Package Includes:</strong><br>',
   '<br> Package included ：<br>': '<br><br><strong>Package Includes:</strong><br>',
   '<br> Package Included: <br> ' : '<br><br><strong>Package Includes:</strong><br>',
   '<br> Pakage Include: <br> ': '<br><br><strong>Package Includes:</strong><br>',
   '<br> • KIT Includes:': '<br><br><strong>Package Includes:</strong><br>',
   ("<br> Package Included：<br> "): '<br><br><strong>Package Includes:</strong><br>',
   '<br>• Package Includes: (optional)': '<br><br><strong>Package Includes:</strong><br>'
}

df['Body HTML'] = df['Body HTML'].replace(package_pattern, regex=True)

indexes_without_package_includes = df[~df['Body HTML'].str.contains('Package Includes:', case=False)].index
print(f"Count_indexes_without_package_includes: {len(indexes_without_package_includes)}")
print(f"indexes_without_package_includes: {indexes_without_package_includes}")


pattern_to_remove = r'\(\d+/\d+[A-Za-z]?\)'

df['Body HTML'] = df['Body HTML'].str.replace(pattern_to_remove, '', regex=True)

indexes_without_package_includes = df[df['Body HTML'].str.contains('<img loading=', case=False)].index
print(f"Count_indexes_with_images: {len(indexes_without_package_includes)}")
print(f"indexes_with_images: {indexes_without_package_includes}")


df['Body HTML'] = df['Body HTML'].str.replace(r'<br>\s*([\w\d]+)\s*(<br>\s*\1\s*)+', r'<br> \1', regex=True)
# Replace the pattern <br> * M<br> * M with "<br> M"
df['Body HTML'] = df['Body HTML'].str.replace(r'<br>\s*\*?\s*([\w\d]+)\s*(<br>\s*\*?\s*\1\s*)+', r'<br> \1', regex=True)

def process_specifications(body_html):
    if 'Tag Size' in body_html:
        extracted_text = body_html.split('Tag Size')[1]
        
        # Additional operations if the pattern exists
        extracted_text = re.sub(r'^\s*\d+\s*%\s*<br>\s*', '', extracted_text)
        extracted_text = re.sub(r'<br>\s*(\d+\s*%)', r': \1', extracted_text)
        
        # Replace patterns like "10.5 cm" with "10 cm" ## modify it to work with 
        pattern = re.compile(r'(\d+\.\d+) cm')
        extracted_text = re.sub(pattern, lambda x: f'{float(x.group(1)):.0f} cm', extracted_text)
        
        pattern1 = re.compile(r'(\d+\.\d+)')
        extracted_text = re.sub(pattern1, lambda x: f'{float(x.group(1)):.0f} cm', extracted_text)

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
                # elif i <= 1:
                #     line = '<br>• ' + line
                formatted_lines.append(line)  
        
        output_text = ' '.join(formatted_lines)
        
        return '<br>• '+'Tag Size |' + output_text
    
# Check if 'Specifications:' is present in 'Body HTML' and apply processing if true

mask = df['Body HTML'].str.contains('Tag Size', case=False, na=False)
# df['Body HTML'] = df['Body HTML'].apply(process_specifications)

df.loc[mask, 'processed_html'] = df.loc[mask, 'Body HTML'].apply(process_specifications)

df.loc[mask, 'Body HTML'] = ( '<strong>Specifications:</strong>' + df.loc[mask, 'processed_html'] )


# Add Bullets to all three sections:
def replace_br_between_keywords_modified(text):
    if pd.notna(text):
        features_index = text.find("<strong>Features:</strong>")
        specs_index = text.find("<strong>Specifications:</strong>")
        package_index = text.find("<strong>Package Includes:</strong>")

        # Independently process each section regardless of whether the previous section was found
        features_to_specs = text[features_index:specs_index].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>') if features_index != -1 else ""
        specs_to_package = text[specs_index:package_index].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>') if specs_index != -1 else ""
        package_to_end = text[package_index:].replace('<br>  •', '<br>').replace('<br> •', '<br>').replace('<br>•', '<br>') if package_index != -1 else ""

        # Combine the processed sections
        text = features_to_specs.replace('<br>', '<br>•') + specs_to_package.replace('<br>', '<br>•') + package_to_end.replace('<br>', '<br>•')

    return text


df['Body HTML'] = df['Body HTML'].apply(replace_br_between_keywords_modified)




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

df['Body HTML'] = df['Body HTML'].str.replace(r'<br>• -', r'<br>• ', regex=True)

df['Body HTML'] = df['Body HTML'].apply(remove_br_before_strong_and_between_br)

df['Body HTML'] = df['Body HTML'].str.replace(r'<br>• L<br>• M', r'<br>• L | M', regex=True)
df['Body HTML'] = df['Body HTML'].str.replace(r'<br>• XL<br>• L', r'<br>• XL | L', regex=True)
df['Body HTML'] = df['Body HTML'].str.replace(r'<br>• \)', '', regex=True)
df['Body HTML'] = df['Body HTML'].str.replace(r'<br>• General Specification', '', regex=True)
df['Body HTML'] = df['Body HTML'].str.replace(r'\(h\)','', regex=True)

specs = {
    'Color<br>• ': 'Color: ',
    'Output<br>• ': 'Output: ',
    'SNR<br>• ': 'SNR: ',
    'Response<br>• ': 'Response: ',
    'BluetoothVersion<br>• ': 'BluetoothVersion: ',
    'Profiles<br>• ': 'Profiles: ',
    'Distance<br>• ': 'Distance: ',
    'Capacity<br>• ': 'Capacity: ',
    'Time<br>• ': 'Time: ',
    'Weight<br>• ': 'Weight: ',
    'Life: <br>• ': 'Life: ',
    'Quality: <br>• ': 'Quality: ',
    'Display: <br>• ': 'Display: ',
    'Technology: <br>• ': 'Technology: ',
    'Functions:<br>• ': 'Functions: ',
    'Size<br>• ': 'Size: ',
    'Brand<br>• ': 'Brand: ',
    'Model<br>• ': 'Model: ',
    'material<br>• ': 'material: ',
    'Magnification<br>• ': 'Magnification: ',
    'diameter<br>• ': 'diameter: ',
    'Waterproof<br>• ': 'Waterproof: ',
    'Frequency<br>• ': 'Frequency: ',
    'Power<br>• ': 'Power: ',
    'Channel<br>• ': 'Channel: ',
    'Battery<br>• ': 'Battery: ',
    'Function<br>• ': 'Function: ',
    'Type<br>• ': 'Type: ',
    'Control<br>• ': 'Control: ',
    'Range<br>• ': 'Range: ',
    'Time <br>•': 'Time: ',
    'Interface<br>• ': 'Interface: ',
    'Ratio<br>• ': 'Ratio: ',
    'x<br>• ': 'x ',
    '<br>• 50': '| 50 cm',
    r'×<br>• ': 'x',
    '<br>• rist': 'rist',
    '<br>• b': 'b',
    '<br>• all': 'all',
    'Name<br>• ': 'Name: ',
    'Material<br>• ': 'Material: ',
    'Sports<br>• ': 'Sports: ',
    r'1X<br>• ': '1 x ',
    r'2X<br>• ': '2 x ',
    'Space<br>• ': 'Space: ',
    'Touch Screen<br>• ': 'Touch Screen: ',
    'Thickness<br>• ': 'Thickness: ',
   'Resolution<br>• ': 'Resolution: ',
   'Icon Position<br>• ': 'Icon Position: ',
   'CPU<br>• ': 'CPU: ',
   'ROM<br>• ': 'ROM: ',
   'RAM<br>• ': 'RAM: ',
   'Operation System<br>• ': 'Operation System: ',
   'USB port<br>• ': 'USB port: ',
   "Voltage<br>• " : 'Voltage: ',
   "Filling<br>• " : 'Filling: ',
   "Version<br>• " : 'Version: '
    
    }
# 323782
###------>>>> df3_result['Body HTML'] .str.contains('1 x<br>• ').sum()


df['Body HTML'] = df['Body HTML'].replace(specs, regex=True)

def replace_text_between_keywords(df, keyword1, keyword2):
    # Count instances before replacement
    instances_before = df['Body HTML'].str.count(f'{keyword1}.*?{keyword2}').sum()
    # Replace text between keywords
    df['Body HTML'] = df['Body HTML'].replace(f'{keyword1}.*?{keyword2}', f'{keyword2}', regex=True)
    # Count instances after replacement
    instances_after = df['Body HTML'].str.count(f'{keyword2}').sum()
    # Print the number of instances replaced
    print(f'*** Replacement took place for {instances_before} instances. ***')

# Define the keywords
keyword1 = '<br><strong>Note:</strong>'
keyword2 = '<br><br><strong>Package Includes:</strong>'

replace_text_between_keywords(df,keyword1,keyword2)


values_to_replace = [
    '<br>• Gift box package',
    '<br>• Caution for the battery:',
    "<br>• Don't over-charge, or over-discharge batteries.",
    "<br>• Don't put it beside the high temperature condition.",
    "<br>• Don't throw it into fire.",
    "<br>• others",
    '<br>• Basic information',
    '<br> Note: The Package Is Not Including Necklace ,Just For Decoration.',
    '<br> Note: Please see the Size Reference to find the correct size.',
    'Note: The Package Is Not Including Necklace ,Just For Decoration.',
    "<br>• Don't throw it into water.",
    "<br>• More details:",
    "<br>• More details",
    "<br>• (This product does not include batteries.)",
    "<br>• #B:",
    "<br>• #A:",
    "<br>• You May Like:",
    "<br>• Manual:",
    "<br>• (Not included battery)",
    "<br>• (Battery not included)",
    "E120 battery life can reach about 15 minutes!<br>• Function"
    
]

# Replace values in the 'Body HTML' column with ''
for value in values_to_replace:
    df['Body HTML'] = df['Body HTML'].str.replace(value, '', case=False)

# Assuming df is your DataFrame
df = df.sort_values(by='Initial_Word', ascending=False)
df = df.reset_index(drop=True)



df['Body HTML'] = df['Body HTML'].str.replace('1x', r'1 x ', regex=True)
df['Body HTML'] = df['Body HTML'].str.replace('1×', r'1 x ', regex=True)
df['Body HTML'] = df['Body HTML'].str.replace('4x', r'4 x ', regex=True)
df['Body HTML'] = df['Body HTML'].str.replace('2x', r'2 x ', regex=True)
df['Body HTML'] = df['Body HTML'].str.replace('3x', r'3 x ', regex=True)
df['Body HTML'] = df['Body HTML'].str.replace('1X', '1 x', regex=True)
df['Body HTML'] = df['Body HTML'].str.replace('1 X', '1 x', regex=True)
df['Body HTML'] = df['Body HTML'].str.replace('2 X', '2 x', regex=True)




skus_to_drop = ['OKPNBXP-7184','OKNPTKT-287832','OKNTNNX-9554-30428','OKKITXB-324184','OPXBNOA','OKIAAKO','OKNKBTA','OKNKIBX-323243'
                , 'OKNXANO','OIPLNAO-6057-311057','OXABLNB-269-984','OOANTTO-20331-287832','OKKOOAN','OKLLPLI-287620-287845',
                'OKNNXOB-18349','OBTNXPT','OKNIOII-17377','OKNPXLA', 'OALTNLP-322452','OXLNINI','OKKOXPA-287832','OITTOXI',
                'OATXIBO','OKNNNKN-7184','OKKKALN','OKTNKNA-318710',
                'OONBTPA-32490','OKNOXBP-7184-293760','OKNXPXT-322291']

dropped_rows = df[df['Variant SKU'].isin(skus_to_drop)].to_dict(orient='records')

# Drop rows where SKU is in the list of SKUs to drop
df = df[~df['Variant SKU'].isin(skus_to_drop)]
df['Body HTML'] = df['Body HTML'].str.replace(r'• \d+>', '•', regex=True)
df['Body HTML'] = df['Body HTML'].str.replace(r'• \*','• ', regex=True)
df['Body HTML'] = df['Body HTML'].str.replace(r': \*',': ', regex=True)

#############################################################################

df['Body HTML'] = df['Body HTML'].apply(lambda x: str(x).split('Tips:')[0] if 'Tips:' in str(x) else x)
df['Body HTML'] = df['Body HTML'].apply(lambda x: str(x).split('Warm Tips')[0] if 'Warm Tips' in str(x) else x)
df['Body HTML'] = df['Body HTML'].apply(lambda x: str(x).split('<br>• More Details: ')[0] if '<br>• More Details: ' in str(x) else x)

# Merge the DataFrames on 'Variant SKU' and replace 'Body HTML' in df2 with 'Body HTML' in df
df3 = pd.merge(df2, df[['Variant SKU', 'Body HTML']], on='Variant SKU',how='inner', suffixes=('_replace', '_original'))

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
df3_result['Tags']= df3_result['Tags'].apply(lambda x: x.replace('not_update_CA', f'DIOWGD_{current_date}') if 'not_update_CA' in x else x)

df3_result['Tags'] =df3_result['Tags'].str.replace(r'color', 'Colour', case=False)

# Apply the modifications directly to the DataFrame
df3_result['Tags'] =df3_result['Tags'].apply(lambda tags: re.sub(r'Shipping_\d+\.\d+,', '', tags, flags=re.IGNORECASE))

# Use regular expression to find the Type value in "Tags" and replace in the "Type" column
df3_result['Type'] = df3_result['Tags'].apply(lambda tags: re.search(r'Type_(.+)', tags, flags=re.IGNORECASE).group(1) if re.search(r'Type_(.+)', tags, flags=re.IGNORECASE) else None)

# Replace values in the Vendor column and store in a dictionary
vendor_mapping = {'idropship': 'GODIAU', 'vidaxl':'GOAUAD', 'wefullfill':'Vibe Geeks', 'bigbuy':'PDBB', 'matterhorn':'GOEFASH','bangood':"Goslash-DIOW"}

df3_result['Vendor'] = df3_result['Vendor'].replace(vendor_mapping)

# Step 3: Replace values in other columns
df3_result['Status'] = df3_result['Status'].replace({'Draft': 'Active'})
df3_result['Published'] = df3_result['Published'].replace({False: True})
df3_result['Published Scope'] = df3_result['Published Scope'].replace({'web': 'global'})

# Convert a set of columns from numeric to text format
columns_to_convert = ['ID', 'Variant Inventory Item ID', 'Variant ID']
df3_result[columns_to_convert] =df3_result[columns_to_convert].astype(str)

replace_text_between_keywords(df3_result, '<strong>Package Includes:</strong>', '<strong>Package Includes:</strong>')
replace_text_between_keywords(df3_result, '<strong>Specifications:</strong>', '<strong>Specifications:</strong>')

replace_text_between_keywords(df3_result, '<br><br><strong>Note:</strong>', '<br><br><strong>Package Includes:</strong>')
replace_text_between_keywords(df3_result, '<br><br><strong>Note:</strong>', '<br><br><strong>Features:</strong>')
replace_text_between_keywords(df3_result, '<strong>Features:</strong>', '<strong>Features:</strong>')
replace_text_between_keywords(df3_result, '<br>• Notice ：<br>•', '<br><br><strong>Package Includes:</strong>')

# df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_br_before_strong_and_between_br)
# df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: str(x).split('<br><strong>Note:</strong>')[0] if '<strong>Note:</strong>' in str(x) else x)
# df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: str(x).split('Note:')[0] if 'Note:' in str(x) else x)
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('(do not have transmitter,receiver,body shell,glow starter and nitro fuel)','')

# ## Adding Package Includes for those were it is missing:
# def process_body_html(row):
#     package_includes_str = '<br><strong>Package Includes:</strong><br> • 1 x '
    
#     if '<br><br><strong>Package Includes:</strong><br>' not in row['Body HTML']:
#         first_three_words = ' '.join(row['Title'].split()[:3]).upper()
#         row['Body HTML'] += package_includes_str + first_three_words

#     return row['Body HTML']



# indexes_without_package_includes = df3_result[~df3_result['Body HTML'].str.contains('Package Includes:', case=False)].index
# print(f"Count_indexes_without_package_includes: {len(indexes_without_package_includes)}")
# print(f"indexes_without_package_includes: {indexes_without_package_includes}")
# Sort by 'Initial_Word'


# df3_result = df3_result[df3_result['Body HTML'].notna() & (df3_result['Body HTML'] != '')]
# df3_result  = df3_result.drop(8)

# # Extract initial word
# df3_result['Initial_Word'] = df3_result['Body HTML'].str.split().str[0]
# # Count occurrences
# df3_result_word_counts = df3_result['Initial_Word'].value_counts()

# # Sort by 'Initial_Word'
# df3_result_sorted = df3_result[['Body HTML','Initial_Word']].sort_values(by='Initial_Word')

# df3_result['Body HTML'] = df3_result.apply(process_body_html, axis=1) 

# df3_result['Body HTML'] =df3_result['Body HTML'].str.replace('<strong<br><strong>','<br><strong>')


# indexes_without_package_includes = df3_result[~df3_result['Body HTML'].str.contains('Package Includes:', case=False)].index
# print(f"Count_indexes_without_package_includes: {len(indexes_without_package_includes)}")
# print(f"indexes_without_package_includes: {indexes_without_package_includes}")

# #df3_result_word_counts.index
# # '<strong>Specifications:</strong><br>•'

# # Define the indices to apply the replacement
# indices_to_replace = [18,33,38,42,43,86,107,128,201,211,229,236,241,262,263,264,275,286,290,291,292,303,305,315,317,329]

# # Define the keywords
# keyword1 = '<br><br><strong>Features:</strong>'
# keyword2 = '<br><br><strong>Package Includes:</strong>'

# # Create a condition based on the 'Body HTML' column
# condition = df3_result['Body HTML'].str.startswith('<strong>Specifications:</strong><br>•')

# # Apply the replacement function to the specified condition
# df3_result.loc[condition, 'Body HTML'] = df3_result.loc[condition, 'Body HTML'].apply(
#     lambda x: re.sub(f'{keyword1}.*?{keyword2}', f'{keyword2}', x, flags=re.DOTALL)
# )


# df3_result_specs=df3_result.loc[df3_result['Body HTML'].str.startswith('<strong>Specifications:</strong><br>•')]




# def syntax_package_includes(text):
#     if pd.notna(text):
#         # Find the starting index of "<strong>Package Includes:</strong>"
#         package_index = text.find("<strong>Package Includes:</strong>")
#         if package_index != -1:
#             # Extract the text that comes after "<strong>Package Includes:</strong>"
#             package_text = text[package_index + len("<strong>Package Includes:</strong>"):]

#             # Apply the replacement only to the extracted package text
#             package_text = re.sub(r'<br>• (.*?)(?:\s*x\s*(\d+|\b(?:one|two|three|four|five|six|seven|eight|nine|ten)\b))', r'<br>• \2 x \1', package_text, flags=re.IGNORECASE)

#             # Replace the original package text with the modified version
#             text = text[:package_index + len("<strong>Package Includes:</strong>")] + package_text

#     return text


# df3_result_specs['Body HTML'] = df3_result_specs['Body HTML'].apply(syntax_package_includes)


# def remove_large_highlights(row):
#     start_pattern = r'<strong>Features:</strong><br>'
#     end_pattern = r'<br><strong>Package Includes:</strong><br>'
    
#     # Extract text between start and end patterns
#     match = re.search(f'{start_pattern}(.*?){end_pattern}', row, re.DOTALL)
#     if match:
#         highlighted_text = match.group(1)
        
#         # Check if there is only '<br><br>' and no other text, and if character count exceeds 1600
#         if len(row) > 1600:
#             # Remove the entire block
#             return row.replace(f'{start_pattern}{highlighted_text}{end_pattern}','<br><strong>Package Includes:</strong><br>')
#     return row


# df3_result['Body HTML'] = df3_result_specs['Body HTML'].apply(remove_large_highlights)

# #df3_result_specs['Body HTML'].str.contains("• Package Includes:").sum()



# df3_result_specs['Body HTML'] =df3_result_specs['Body HTML'].apply(remove_br_before_strong_and_between_br)



# # Define the pattern transformation function
# def transform_pattern(text):
#     # Check if the text contains '<strong>Package Includes:</strong>'
#     if '<strong>Package Includes:</strong>' in text:
#         # Split the text at '<strong>Package Includes:</strong>'
#         after_package_includes = text.split('<strong>Package Includes:</strong>')[1]
#         before_package_includes = text.split('<strong>Package Includes:</strong>')[0]
#         # Apply the pattern transformation to the part after '<strong>Package Includes:</strong>'
#         transformed_part = re.sub(r'(\d+[Xx])', r'<br>• \1', after_package_includes)
#         # Combine the modified part with the '<strong>Package Includes:</strong>' prefix
#         return before_package_includes + f'<strong>Package Includes:</strong>{transformed_part}'
#     else:
#         # Return the original text if '<strong>Package Includes:</strong>' is not found
#         return text

# # Apply the function to the 'Body HTML' column
# df3_result_specs['Body HTML'] = df3_result_specs['Body HTML'].apply(transform_pattern)


# df3_result_specs['Body HTML'] =df3_result_specs['Body HTML'].apply(remove_br_before_strong_and_between_br)

# df3_result=df3_result_specs

# Assuming df3_result is your DataFrame
max_length = 1600
#df3_result = df3_result[df3_result['Body HTML'].str.len() <= max_length].reset_index(drop=True)

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('• x ','• ')


df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: re.sub(r'• \d+\)|• \d+\.', '• ', x))



def reformat_html_body(body_html):
    # Define patterns to identify each section
    features_pattern = r'(?s)(<strong>Features:</strong>.*?)(?=<strong>|$)'
    specs_pattern = r'(?s)(<strong>Specifications:</strong>.*?)(?=<strong>|$)'
    package_includes_pattern = r'(?s)(<strong>Package Includes:</strong>.*?)(?=<strong>|$)'

    # Find each section using regex
    features = re.search(features_pattern, body_html)
    specifications = re.search(specs_pattern, body_html)
    package_includes = re.search(package_includes_pattern, body_html)

    # Extract text if the section is found, else use an empty string
    features_text = features.group(0) if features else ''
    specs_text = specifications.group(0) if specifications else ''
    package_includes_text = package_includes.group(0) if package_includes else ''

    # Combine in the desired order
    combined_text = f'{features_text}{specs_text}{package_includes_text}' + '<br>'
    return combined_text

# Apply the reformatting function to the 'Body HTML' column
df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: reformat_html_body(x) if pd.notnull(x) else x)


def trim_section_to_limit(section_html, char_limit):
    """
    Revised function to trim the HTML text of a section (either Features or Specifications) 
    while preserving <br><br> tags, ensuring the total character count 
    does not exceed the specified limit.

    :param section_html: HTML text of the section to be trimmed.
    :param char_limit: Character limit to adhere to.
    :return: Trimmed HTML text of the section.
    """
    # Replace <br><br> with a placeholder to preserve them during splitting
    section_html_placeholder = section_html.replace('<br><br>', '[[BR_BR]]')

    # Split the section into lines based on <br>•
    lines = section_html_placeholder.split('<br>•')
    
    # Reconstruct the section, removing lines from the end until it's within the character limit
    for i in range(len(lines), 0, -1):
        trimmed_text = '<br>•'.join(lines[:i])

        # Replace the placeholder back to <br><br>
        trimmed_text_with_br = trimmed_text.replace('[[BR_BR]]', '<br><br>')

        if len(trimmed_text_with_br) <= char_limit:
            return trimmed_text_with_br

    return ''  # Return empty string if it's not possible to trim within the limit

def trim_text_accordingly(html_text, char_limit=1600):
    # Extract each section
    features = re.search(r'(?s)(<strong>Features:</strong>.*?)(?=<strong>|$)', html_text)
    specifications = re.search(r'(?s)(<strong>Specifications:</strong>.*?)(?=<strong>|$)', html_text)
    package_includes = re.search(r'(?s)(<br><br><strong>Package Includes:</strong>.*?)(?=<strong>|$)', html_text)

    # Get text for each section, if found
    features_text = features.group(0) if features else ''
    specs_text = specifications.group(0) if specifications else ''
    package_includes_text = package_includes.group(0) if package_includes else ''

    # Determine which section to trim (Features or Specifications)
    if features:
        primary_section_text = features_text
        secondary_section_text = specs_text + package_includes_text
    else:
        primary_section_text = specs_text
        secondary_section_text = package_includes_text

    # Trim the primary section if the total length exceeds the limit
    if len(primary_section_text) + len(secondary_section_text) > char_limit:
        trimmed_primary_text = trim_section_to_limit(primary_section_text, char_limit - len(secondary_section_text))
    else:
        trimmed_primary_text = primary_section_text

    # Reconstruct the HTML text with the trimmed primary section
    if features:
        combined_text = f'{trimmed_primary_text}<br><br>{specs_text}{package_includes_text}'
    else:
        combined_text = f'{trimmed_primary_text}{package_includes_text}'

    return combined_text


# Apply the trimming function to the 'Reformatted Body HTML' column
df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: trim_text_accordingly(x) if pd.notnull(x) else x)


df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('<br><br><br><br>','<br><br>')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('<br><br><br>','<br><br>')

 





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
        
        final_text = final_text = [sentence.strip().lower() if sentence.lower().strip() in custom_stop_words 
               else sentence.strip().title() if (sentence.count(' ') <= 5  and sentence.lower().strip() not in custom_stop_words)  or ':' in sentence
               else sentence.strip().capitalize() 
               for sentence in sentences]

        


        final_text = ' '.join(final_text)

            
        
        # Remove extra full stops after punctuation signs
        final_text = re.sub(r'(?<=[.!?])\s*\.', '', final_text)

        # Replace the original highlighted text with the modified version
        text = text[:start_pos] + final_text + text[end_pos:]

    return text


def capitalize_sentences_v2(text, start_pattern, end_pattern):
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
        
        final_text = [
            sentence.strip().lower() if sentence.lower().strip() in custom_stop_words 
            else sentence.strip().title() if (sentence.count(' ') <= 5 and sentence.lower().strip() not in custom_stop_words)
            else sentence.strip().capitalize() 
            for sentence in sentences
        ]

        # Additional functionality for • and :
        for i, sentence in enumerate(final_text):
            if '•' in sentence:
                # Capitalize the text between • and :
                parts = sentence.split('•')
                if len(parts) > 1:
                    parts[1] = parts[1].strip().capitalize()
                final_text[i] = '•'.join(parts)
            
            if ':' in sentence:
                # Capitalize the text between : and add a space after :
                parts = sentence.split(':')
                if len(parts) > 1:
                    parts[1] = parts[1].strip().capitalize()
                final_text[i] = ': '.join(parts)

        final_text = ' '.join(final_text)

        # Remove extra full stops after punctuation signs
        final_text = re.sub(r'(?<=[.!?])\s*\.', '', final_text)

        # Replace the original highlighted text with the modified version
        text = text[:start_pos] + final_text + text[end_pos:]

    return text



start_pattern = r'<strong>Features:</strong><br>'
end_pattern = r'<br><br><strong>Specifications:</strong><br>'

df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: capitalize_sentences_v2(x, start_pattern, end_pattern))

start_pattern = r'<br><br><strong>Specifications:</strong><br>'
end_pattern =  r'<br><br><strong>Package Includes:</strong><br>'
df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: capitalize_sentences(x, start_pattern, end_pattern))


df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('xxxxxl','XXXXXL')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('xxxxl','XXXXL')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('xxxl','XXXL')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('xxl','XXL')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('xl','XL')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('Xl','XL')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('m,l,s,' , 'M,L,S,')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(',l,' , ',L,')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(',m,' , ',M,')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(',s,' , ',S,')

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('l,s,' , 'L,S,')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('Color' , 'Colour')

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('color' , 'Colour', case = False)

## Adding Package Includes for those were it is missing:
def process_body_html(row):
    package_includes_str = '<br><strong>Package Includes:</strong><br> • 1 x '
    
    if '<strong>Package Includes:</strong>' not in row['Body HTML']:
        first_three_words = ' '.join(row['Title'].split()[:3]).upper()
        row['Body HTML'] += package_includes_str + first_three_words

    return row['Body HTML']



def reformat_package_includes_section(html_text):
    package_includes_pattern = r'(?s)(<strong>Package Includes:</strong>.*?)(?=<strong>|$)'
    package_includes_section = re.search(package_includes_pattern, html_text)

    if not package_includes_section:
        return html_text

    modified_package_includes = []
    lines = package_includes_section.group(0).split('<br>•')

    for line in lines:
        # Adjusted regex to match patterns like 'USBx1' or 'USB x 1'
        modified_line = re.sub(r'([\w\s]+?)\s*x\s*(\d+)', r'\2 x \1', line)
        modified_package_includes.append(modified_line)

    reconstructed_package_includes = '<br>• '.join(modified_package_includes)
    modified_html_text = html_text.replace(package_includes_section.group(0), reconstructed_package_includes)
    
    return modified_html_text

df3_result['Body HTML'] = df3_result['Body HTML'].apply(reformat_package_includes_section)

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('*', '')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace(',', ', ')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('+', ' + ')


condition = df3_result['Variant SKU'] == 'OKKPONT'
df3_result.loc[condition, 'Body HTML'] = '<strong>Specifications:</strong>' + df3_result.loc[condition, 'Body HTML'].str.split('<strong>Specifications:</strong>').str[1].fillna('')


condition = df3_result['Variant SKU'] == 'OKTKTTX'
df3_result.loc[condition, 'Body HTML'] = '<strong>Specifications:</strong>' + df3_result.loc[condition, 'Body HTML'].str.split('<strong>Specifications:</strong>').str[1].fillna('')


condition = df3_result['Variant SKU'] == 'OKKOOKL-311057'
df3_result.loc[condition, 'Body HTML'] = '<strong>Specifications:</strong>' + df3_result.loc[condition, 'Body HTML'].str.split('<strong>Specifications:</strong>').str[1].fillna('')


condition = df3_result['Variant SKU'] == 'OPPTNXT'
df3_result.loc[condition, 'Body HTML'] = df3_result.loc[condition, 'Body HTML'].str.split('On The Screen').str[0].fillna('') + 'On The Screen <br>'

condition = df3_result['Variant SKU'] == 'OKKPLPX-287638'
df3_result.loc[condition, 'Body HTML'] = df3_result.loc[condition, 'Body HTML'].str.split('• Press and hold').str[0].fillna('') 

condition = df3_result['Variant SKU'] == 'OKKPLNB-6726'
df3_result.loc[condition, 'Body HTML'] = df3_result.loc[condition, 'Body HTML'].str.split('• Heated Jacket With High').str[0].fillna('') 



condition = df3_result['Variant SKU'] == 'OKNNTPT-2481-292556'
df3_result.loc[condition, 'Body HTML'] = '<strong>Specifications:</strong>' + df3_result.loc[condition, 'Body HTML'].str.split('<strong>Specifications:</strong>').str[1].fillna('')


condition = df3_result['Variant SKU'] == 'OPPTNXT'
df3_result.loc[condition, 'Body HTML'] = '<strong>Specifications:</strong>' + df3_result.loc[condition, 'Body HTML'].str.split('<strong>Specifications:</strong>').str[1].fillna('')



condition = df3_result['Variant SKU'] == 'XBBBAOK-287845'
df3_result.loc[condition, 'Body HTML'] = '<strong>Features:</strong><br>• ' + df3_result.loc[condition, 'Body HTML'].str.split('<strong>Features:</strong>').str[1].fillna('')

condition = df3_result['Variant SKU'] == 'OKKPLLI-287830'
df3_result.loc[condition, 'Body HTML'] = '<strong>Specifications:</strong>' + df3_result.loc[condition, 'Body HTML'].str.split('<strong>Specifications:</strong>').str[1].fillna('')


condition = df3_result['Variant SKU'] == 'OKNBLBK-322049'
df3_result.loc[condition, 'Body HTML'] =  df3_result.loc[condition, 'Body HTML'].str.replace('<strong>Package Includes:</strong><br>','<br><br><strong>Package Includes:</strong><br>')



#OKNBLBK-322049



spec={ r': 1 x ':  '<br>• 1 x ' ,
      'Ieee' : 'IEEE',
      '• Function <br>' :'',
      r' 1 x ':  '<br>• 1 x '
      }

df3_result['Body HTML'] = df3_result['Body HTML'].replace(spec, regex=True)

df3_result['Body HTML'] = df3_result.apply(process_body_html, axis=1) 

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
        patterns_between_br = ['<br>•<br>', '<br>• <br>','<br>•  <br>','<br> •<br>','<br> • <br>']
        for pattern in patterns_between_br:
            while pattern in text:
                text = text.replace(pattern, '<br>')
        
        # Remove all occurrences of <br>• at the end
        text = re.sub(r'<br>•\s*$', '', text)
    
    return text

df3_result['Body HTML'] = df3_result['Body HTML'].apply(remove_br_before_strong_and_between_br)



df3_sorted, df3_counts = extract_and_return_word_counts(df3_result)

df3_result = df3_result[column_order]

def make_tags_lowercase(text):
    patterns = [r'<Br>', r'<Strong>', r'</Strong>']

    for pattern in patterns:
        text = re.sub(re.escape(pattern), pattern.lower(), text)

    return text

df3_result['Body HTML'] = df3_result['Body HTML'].apply(make_tags_lowercase)


# Check for 'Eu', 'Us', 'Uk', and 'plug' in Title column
eu_us_uk_condition_title = df3_result['Title'].str.contains(r'\b(Eu|Us|Uk)\b', case=False)
df3_result.loc[eu_us_uk_condition_title, 'Command'] = 'DELETE'

# Check for 'Eu', 'Us', 'Uk', and 'plug' in Body HTML column
eu_us_uk_condition_body = df3_result['Body HTML'].str.contains(r'\b(Eu|Us|Uk)\b', case=False)
df3_result.loc[eu_us_uk_condition_body, 'Command'] = 'DELETE'

# Print indices where the conditions are met
print("Indices to delete:")
print(df3_result[df3_result['Command'] == 'DELETE'].index)

# Function to capitalize the first letter of each word after |
def capitalize_after_pipe(text):
    parts = text.split('|')
    updated_parts = []
    for part in parts:
        if part.strip():
            # Preserve case for the rest of the text and capitalize only the first letter
            updated_parts.append(part.strip()[0].capitalize() + part.strip()[1:])
        else:
            updated_parts.append(part.strip())
    return '|'.join(updated_parts)

df3_result['Body HTML'] = df3_result['Body HTML'].apply(capitalize_after_pipe)

df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: re.sub(r'(\d+)\s*\.\s*(\d+)', r'\1.\2', x))

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('Sleeve length', 'Sleeve Length')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('Bottom length', 'Bottom Length')

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('•  ','• ')

df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: re.sub(r'<strong>Package Includes:</strong>', '<strong>Package Includes:</strong> (as per your selection)', x) if '• tag size|' in x.lower() else x)

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('• •','•')


# Initialize a list to keep track of indices where replacements occur
indices_with_replacements = []

# Apply the function to 'Body HTML' and 'Title' columns
apply_removal_and_track_indices(df3_result, 'Body HTML', words_to_remove)
apply_removal_and_track_indices(df3_result, 'Title', words_to_remove)

# Now print the indices where replacements occurred
print(sorted(set(indices_with_replacements)))

# Apply the trimming function to the 'Reformatted Body HTML' column
df3_result['Body HTML'] = df3_result['Body HTML'].apply(lambda x: trim_text_accordingly(x) if pd.notnull(x) else x)
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('<br><br><br><br>','<br><br>')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('<br><br><br>','<br><br>')

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('<br>• 1 x <br>','<br>')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('bo x','box')

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('<br>• 150 x <br>','<br>')
df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('<br>• 200 x <br>','<br>')

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('<br>• Only the above package content,  and other products are not included.','')

df3_result['Body HTML'] = df3_result['Body HTML'].str.replace('<br><br><br><br>','<br><br>')


# ### HTML: 

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