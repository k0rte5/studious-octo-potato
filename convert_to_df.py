# -*- coding: utf-8 -*-
# Importing libraries
import pandas as pd
import openpyxl

# Set the variables for reading the Excel file, can be changed as needed
filename= 'ИХТиПЭ-очно-2курс.xlsx'                                                            # Filename for the downloaded Excel file
student_group = 'ХПУ-124'                                                                     # Student group which is also a sheet name in the Excel file

# Read the Excel file
df = pd.read_excel(filename, sheet_name=student_group, usecols='B:N', skiprows=15, nrows=102) # Schedule is always in range B16:N117 

print(df.iloc[6])  # Print the first row to verify correct reading