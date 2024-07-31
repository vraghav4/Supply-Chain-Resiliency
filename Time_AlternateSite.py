# -*- coding: utf-8 -*-
"""
Created on Wed Jul 31 10:59:26 2024

@author: 267447
"""

import pandas as pd

# Specify the path to your Excel file
excel_file = 'C:/Users/267447/Documents/Python Scripts/SC Resiliency/Test 07132024.xlsx'

# Read the Excel file into a dataframe
df = pd.read_excel(excel_file)

# Print the dataframe
print(df)

#calculate available capacity at each node
df.loc [:, 14] = df.iloc[:,10] * (1- df.iloc[:,13]) / df.iloc[:,13]

#bound lowest value to zero
df.loc[:,14] = df.loc[:,14].clip(lower=0)