# -*- coding: utf-8 -*-
"""
Created on Thu Aug  1 15:04:16 2024

@author: venim
"""


import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog


def select_file():
    global excel_file
    excel_file = filedialog.askopenfilename(filetypes = [('Excel files', '*xlsx')])
    if excel_file:
       try:
           df = pd.read_excel(excel_file)
           # Do something with the dataframe
           tk.messagebox.showinfo("Success", "File uploaded successfully!")
       except Exception as e:
           tk.messagebox.showerror("Error", f"Error uploading file: {e}")
    


def calculate_time_to_address(excel_file):
     


    # Specify the path to your Excel file
    # C:\Users\venim\Documents\Areas\Supply Chain Analytics\Analysis
    # excel_file = 'C:/Users/venim/Documents/Areas/Supply Chain Analytics/Analysis/Test 07062024.xlsx'
    
    # Read the Excel file into a dataframe
    df = pd.read_excel(excel_file)
    
   
    #calculate available capacity at each node
    df.loc [:, 'Available Capacity'] = df.loc[:,'Volume + Demand Index'] * (1- df.loc[:,'Current utilization']) / df.loc[:,'Current utilization']
    
    #bound lowest value to zero
    df.loc[:,'Available Capacity'] = df.loc[:,'Available Capacity'].clip(lower=0)
    
    #calculate total volume for each material
    TotalVolume = df.groupby('Pdm Matl Nbr')['Volume + Demand Index'].sum().reset_index()
    
    
    #Calculate Capacity at Alternate Sites
    df['Vol at Alternate Site'] = df.groupby('Pdm Matl Nbr')['Available Capacity'].transform('sum') - df['Available Capacity']
    
    #Calculate Average Lead time for all the Alternate Sites
    df['Average Lead Time'] = df.groupby('Pdm Matl Nbr')['Lead time'].transform('mean')
    
    #Calculate time to bring in Alternate Site when node goes down
    #Assumption baked in, 7 extra days on top of the lead time to bring the alternate site to deliver the extra volume
    df['Time to bring in Alternate Site'] = df.apply(lambda x: x['Average Lead Time'] + 7 if x['Vol at Alternate Site'] >= x['Volume + Demand Index'] else 0, axis = 1 )
    
    #Days in Transit when node goes down
    #Assumptions, Delivery Frequency = 2 times a week
    df['Days in Transit'] = df['Lead time'] *2 / (7)
    
    #Time to Survive = Days in Transit plus T1 inventory, T2 inventory ignored due to assumption of decimation of the node an no access to retrieve the staged inventory
    #T1 inventory is the lead time, i.e. T1 facility keeps lead time worth of product in its inventory
    df['Time to Survive'] = df['Days in Transit'] + df['Lead time']
    
    #Time to Recover, assuming same MCO facility can be used to bridge the demand gap. 
    df['Time to Recover'] = df['Time to bring in Alternate Site']
    
    #Time to Recover - Time to Survive, i.e. The gap to bridge in the migitation effort 
    df['TTR - TTS'] = np.clip(df['Time to Recover'] - df['Time to Survive'], a_min=0, a_max=None)

    return df   


def calculate_and_display():
    """Calculates the time to recover and displays it in the GUI."""
    if not excel_file:
        result_label.config(text="Please select an Excel file")
        return

    try:
        time_to_address = calculate_time_to_address(excel_file)
        result_label.config(text=f"TTR - TTS:\n\n {time_to_address.to_string(index=False, columns = ['Pdm Matl Nbr', 'Vendor Code','Second Suffix', 'TTR - TTS' ])}")
    except Exception as e:
        result_label.config(text=f"Error: {str(e)}")

#Creating the GUI

root = tk.Tk()
root.title("Time to Address Calculator")

file_button = tk.Button(root, text="Select Excel File", command=select_file)
file_button.pack()

result_label = tk.Label(root, text="")
result_label.pack()

calculate_button = tk.Button(root, text="Calculate", command=calculate_and_display)
calculate_button.pack()


root.mainloop()