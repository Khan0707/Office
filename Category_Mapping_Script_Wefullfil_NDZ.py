# -*- coding: utf-8 -*-
"""
Created on Mon Jan 29 09:52:02 2024

@author: aaqib
"""

import pandas as pd
from difflib import SequenceMatcher

# Specify the Excel file path
excel_file_path = r"D:\BackendData\DSZ_Wefullfil_CategortyMapping_File.xlsx"

# Read the first sheet into DataFrame df1
df1 = pd.read_excel(excel_file_path, sheet_name='Wefullfil Categories')

# Read the second sheet into DataFrame df2
df2 = pd.read_excel(excel_file_path, sheet_name='Dropshipzone Categories')


# Function to calculate similarity ratio
def similarity_ratio(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()


column_df1 = 'Wefullfil_Category'
column_df2 = 'Dsz_Category'

# # Create a new column in df1 to store the mapped values
df1['Mapped_Category'] = df1[column_df1].apply(lambda x: max(df2[column_df2], key=lambda y: similarity_ratio(x, y)) if pd.notna(x) else None)

######################

# import pandas as pd
# from difflib import SequenceMatcher

# # Assuming df1 and df2 are your DataFrames
# # Replace these with the actual column names from your DataFrames
# column_df1 = 'Wefullfil_Category'
# column_df2 = 'Dsz_Category'

# # Function to calculate similarity ratio
# def similarity_ratio(a, b):
#     return SequenceMatcher(None, str(a), str(b)).ratio()

# # List of threshold values
# thresholds = [0.9, 0.8, 0.7, 0.6, 0.5,0.4,0.3]

# # Perform iterations for each threshold
# for i, threshold in enumerate(thresholds):
#     # Create a new column for each iteration
#     new_column_name = f'Mapped_Category_{i+1}_Threshold_{threshold}'
    
#     # Update the new column, excluding previously mapped values
#     df1[new_column_name] = df1[column_df1].apply(lambda x: max(
#         (y for y in df2[column_df2] if similarity_ratio(x, y) > threshold),
#         default=None,
#         key=lambda y: similarity_ratio(x, y)
#     ) if pd.notna(x) else None)

# # Print the final mapped DataFrame
# print(df1)
