''' Finds and lists duplicate values from an Excel spreadsheet, created by Jeremy Reynolds.
    Input file path (Line 8), and column name to use (Line 9).
    1,000,000 values takes ~30 seconds(linear), uses <250MB of memory, and <15% CPU on an average system. '''

import pandas as pd
import os

path = 'test.xlsx'                                  # Input file, Syntax:/Users/Username/Desktop/test.xlsx
column = 'from_user'                                # The column name for duplicate searching
outFile = 'Duplicates.csv'                          # Save file name

xlsx = pd.ExcelFile(path)
df = pd.read_excel(xlsx)                            # Use (xlsx, 'Sheet1') to indicate sheet number

df['is_duplicated'] = df.duplicated(column, keep=False)
''' Creates new column 'is_duplicated' and lists duplicates as True or False
    keep=False keeps all duplicate values including the first
    Use keep='first' to exclude first duplicate value '''

df = df.loc[df['is_duplicated'] == True]            # Creates a new column with rows only where 'is_duplicated' = True

df = df.drop_duplicates(column, keep='first')       # Overwrites df with the first duplicates only (Comment line for all duplicates)

total_dup = df['is_duplicated'].sum()               # Calculates total duplicates and prints
print('Duplicates:  ',total_dup

df.drop('is_duplicated', axis=1, inplace=True)      # Removes 'is_duplicated' row

df.to_csv(outFile, index=False, encoding='utf-8')   # index=False removes column with the original index value

print('File Saved at',os.getcwd())
