#!/usr/bin/env python
# coding: utf-8

# In[1]:


#!/usr/bin/env python
# coding: utf-8

import os
import pandas as pd
import numpy as np
from dateutil import parser
from xlsxwriter import Workbook

def print_working_directory():
    print(os.getcwd())



# In[2]:


def load_and_inspect_data(file_path):
    # Loading the phishing dataset
    data = pd.read_excel(file_path, sheet_name='Incident Dataset')
    print("First few rows of the dataset:")
    print(data.head())
    print("\nSummary of missing values in each column:")
    print(data.isnull().sum())
    print("\nData types of each column:")
    print(data.dtypes)
    return data



# In[3]:


def clean_incident_data(data):
    # Cleaning null columns and rows
    data.drop(data.columns[0], axis=1, inplace=True)
    new_header = data.iloc[1]
    data2 = data[2:]
    data2.columns = new_header
    
    # Processing Employee column
    data2_copy = data2.copy()
    data3 = data2_copy.sort_values(by='Employee')
    data3['Employee'] = data3['Employee'].fillna(0).astype(int)
    max_serial_a = data3['Employee'].max()
    zeros_mask_a = data3['Employee'] == 0
    data3.loc[zeros_mask_a, 'Employee'] = np.arange(max_serial_a + 1, max_serial_a + 1 + zeros_mask_a.sum())
    
    # Convert columns to numeric
    data3['Clicked on Link?'] = data3['Clicked on Link?'].apply(lambda x: 1 if x == 'Yes' else 0)
    data3['# of Fails in Past 6 Months'] = data3['# of Fails in Past 6 Months'].astype(int)
    data3['# of Fails in Past 13 Months'] = data3['# of Fails in Past 13 Months'].astype(int)
    columns_to_fill = data3.columns[2:15]
    data3[columns_to_fill] = data3[columns_to_fill].fillna(0)
    
    # Clean Division column
    data3['Division'] = data3['Division'].replace(' ', 'Unknown').fillna('Unknown')
    data3['Division'] = data3['Division'].str.strip().str.upper()
    data3['HighRisk'] = data3['# of Fails in Past 13 Months'].apply(lambda x: 'HIGH' if x >= 2 else 'LOW')
    data3['Device Type Used'] = data3['Device Type Used'].str.upper()
    data3.drop_duplicates(inplace=True)
    data3['Clicked on Link?'] = data3['# of Fails in Past 13 Months'].apply(lambda x: 1 if x > 0 else 0)
    
    # Summary after cleaning
    print("\nSummary after data cleaning:")
    print(data3.head())
    print("\nSummary of missing values in each column:")
    print(data3.isnull().sum())
    print("\nRows with missing values:")
    rows_with_missing_i = data3[data3.isnull().any(axis=1)]
    print(rows_with_missing_i)
    print("\nData types of each column:")
    print(data3.dtypes)
    
    return data3



# In[4]:


def clean_employee_data(file_path):
    # Now clean Employee sheet
    dfemp = pd.read_excel(file_path, sheet_name='Employee Data')
    dfemp.drop(dfemp.columns[0], axis=1, inplace=True)
    emp_header = dfemp.iloc[1]
    dfemp = dfemp[2:]
    dfemp.columns = emp_header
    
    dfemp2 = dfemp.copy()
    
    # Function to convert dates to standard format with error handling
    def convert_to_standard_date(date_str):
        try:
            return pd.to_datetime(date_str).strftime('%Y-%m-%d')
        except ValueError:
            parts = date_str.split('/')
            month = int(parts[0])
            day = int(parts[1])
            year = int(parts[2])
            if month == 2 and day > 28:
                day = 28
            elif month in [4, 6, 9, 11] and day > 30:
                day = 30
            else:
                day = 31
            corrected_date = f'{month}/{day}/{year}'
            return pd.to_datetime(corrected_date).strftime('%Y-%m-%d')

    # Apply the conversion function to the date_column
    dfemp2['Hire Date'] = dfemp2['Hire Date'].apply(convert_to_standard_date)
    
    dfemp2['Division'] = dfemp2['Division'].replace(' ', 'Unknown')
    dfemp2['Employee Region'] = dfemp2['Employee Region'].replace(' ', 'Unknown')
    dfemp2['Employee'] = dfemp2['Employee'].fillna(0).astype(int)
    max_serial = dfemp2['Employee'].max()
    zeros_mask = dfemp2['Employee'] == 0
    dfemp2.loc[zeros_mask, 'Employee'] = np.arange(max_serial + 1, max_serial + 1 + zeros_mask.sum())
    dfemp2['Division'] = dfemp2['Division'].fillna('Unknown')
    dfemp2['Employee Region'] = dfemp2['Employee Region'].fillna('Unknown')
    dfemp2['Hire Date'] = pd.to_datetime(dfemp2['Hire Date'])
    
    # Summary after cleaning
    print("\nSummary of missing values in each column:")
    print(dfemp2.isnull().sum())
    print("\nRows with missing values:")
    rows_with_missing3 = dfemp2[dfemp2.isnull().any(axis=1)]
    print(rows_with_missing3)
    print("\nData types of each column:")
    print(dfemp2.dtypes)
    
    return dfemp2



# In[5]:


def save_cleaned_data(data3, dfemp2, cleaned_file_path):
    with pd.ExcelWriter(cleaned_file_path, engine='xlsxwriter') as writer:
        data3.to_excel(writer, sheet_name='Incident', index=False)
        dfemp2.to_excel(writer, sheet_name='Employee', index=False)
    print(f"\nCleaned data saved to {cleaned_file_path}")

    
def main():
    #modify both of the path according to your system
    file_path = '/Users/amanpreet/Downloads/IAD_Sample_Data_Set_Project.xlsx' 
    cleaned_file_path = '/Users/amanpreet/Downloads/cleaned_phishing_test_results.xlsx'
    
    print_working_directory()
    
    # Incident data processing
    data = load_and_inspect_data(file_path)
    data3 = clean_incident_data(data)
    
    # Employee data processing
    dfemp2 = clean_employee_data(file_path)
    
    # Save cleaned data
    save_cleaned_data(data3, dfemp2, cleaned_file_path)

if __name__ == "__main__":
    main()

