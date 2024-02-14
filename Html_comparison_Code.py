# -*- coding: utf-8 -*-
"""
Created on Wed Jan 17 07:54:49 2024

@author: aaqib
"""
import pandas as pd
from bs4 import BeautifulSoup

output_file_path = r"D:\BackendData\Aliexpress\Aliexpress_AI_Titles_Desc.xlsx"
df3_result = pd.read_excel(output_file_path)

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
    file.write('    background-color: #e6f7ff; /* Light blue background for Body HTML */\n')
    file.write('    padding: 10px;\n')
    file.write('    border-radius: 5px;\n')
    file.write('    margin-bottom: 15px;\n')
    file.write('}\n')
    file.write('.ai-opt-desc {\n')
    file.write('    font-size: normal;\n')
    file.write('    text-align: justify;\n')
    file.write('    background-color: #f2e6ff; /* Light purple background for AI_Optimised_Desc */\n')
    file.write('    padding: 10px;\n')
    file.write('    border-radius: 5px;\n')
    file.write('    margin-bottom: 15px;\n')
    file.write('}\n')
    file.write('</style>\n')
    file.write('</head>\n<body>\n')

    # Iterate through rows of the DataFrame
    for index, row in df3_result.iterrows():
        variant_sku = row['Variant SKU']
        title = row['Title']
        ai_opt_title = row['AI_Optimised_Title']
        body_html = row['Body HTML']
        ai_optimised_desc = row['AI_Optimised_Desc']

        # Remove images from Body HTML
        soup_body_html = BeautifulSoup(body_html, 'html.parser')
        for img_tag in soup_body_html.find_all('img'):
            img_tag.decompose()
        body_html_no_images = str(soup_body_html)

        # Remove images from AI_Optimised_Desc
        soup_ai_opt_desc = BeautifulSoup(ai_optimised_desc, 'html.parser')
        for img_tag in soup_ai_opt_desc.find_all('img'):
            img_tag.decompose()
        ai_opt_desc_no_images = str(soup_ai_opt_desc)

        # Write the sections to the HTML file with styling
        file.write(f'<main>\n')
        file.write(f'    <h3 style="font-size: larger;">S.No: {index} | Variant SKU: {variant_sku}</h3>\n')
        file.write(f'    <h2>Title: {title}</h2>\n')

        # Section for Body HTML
        file.write(f'    <div class="body-html">\n')
        file.write(f'        <h4>Body HTML</h4>\n')
        file.write(f'        <p>{body_html_no_images}</p>\n')
        file.write(f'    </div>\n')

        # Section for AI_Optimised_Desc
        file.write(f'    <div class="ai-opt-desc">\n')
        file.write(f'        <h4>{ai_opt_title}</h4>\n')
        file.write(f'        <p>{ai_opt_desc_no_images}</p>\n')
        file.write(f'    </div>\n')

        file.write(f'</main>\n\n')

    file.write('</body>\n</html>')