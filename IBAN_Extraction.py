# -*- coding: utf-8 -*-
"""
Created on Mon Mar  6 11:15:32 2023

@author: Chaudhary.Ar
"""

import streamlit as st
import pandas as pd
from schwifty import IBAN
import io
import xlsxwriter


st.write("Step 1: Download Excel Template and Fill the IBAN column")

df = pd.DataFrame(columns=['IBAN'])

# buffer to use for excel writer
buffer = io.BytesIO()

with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
# Write each dataframe to a different worksheet.
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    # Close the Pandas Excel writer and output the Excel file to the buffer
    writer.close()

    download2 = st.download_button(
        label="Download Excel Template",
        data=buffer,
        file_name='IBAN_Template.xlsx',
        mime='application/vnd.ms-excel'
    )
    

st.write("------------------------------------------------------------------------------------")


st.write("Step 2: Upload Filled Template")

excel_file = st.file_uploader("Select your IBAN File to start extraction (Kindly ensure that only one column is provided and is named as 'IBAN')")  

if excel_file is not None:

    df = pd.read_excel(excel_file, dtype = object)

    new = df["IBAN"].astype(str)
    iban1 = []
    bic = []
    account = []
    bank_code = []
    branch_code = []
    country = []
    bank_name = []
    is_valid = []
    
        
    #MAPPING TO THE MAIN BANK SHEET
    #Bank code is Bank Key BANKL
    #Country_code is BANKS
    #Account Number is BANKN
    
    for i in range(len(new)):
        try:
            iban = IBAN(new[i])
            ib = iban.compact 
            cc = iban.country_code 
            bc = iban.bank_code 
            brc = iban.branch_code
            ac = iban.account_code
            bn = iban.bank_name
            bicc = iban.bic
            iv = iban.is_valid
            
            iban1.append(ib)
            account.append(ac)
            bank_code.append(bc)
            country.append(cc)
            branch_code.append(brc)
            bank_name.append(bn)
            is_valid.append(iv)
            bic.append(bicc)
            
            
        except:
            ib = ''
            cc = '' 
            bc = '' 
            ac = ''
            brc = ''
            bn = ''
            iv = ''
            bicc = ''
            
            iban1.append(ib)
            account.append(ac)
            bank_code.append(bc)
            country.append(cc)
            branch_code.append(brc)
            bank_name.append(bn)
            is_valid.append(iv)
            bic.append(bicc)
            
            
    #data = {'IBAN': iban1, 'Account Number': account, 'Bank_code': bank_code, 'Country_code': country, 'branch_code': branch_code, 'Bank_name': bank_name, 'BIC': bic ,'IBAN_is_valid': is_valid}
    data = {'IBAN': iban1, 'BANKN': account, 'BANKL': bank_code, 'Branch_code': branch_code, 'BANKS': country, 'Bank_name': bank_name, 'BIC': bic ,'IBAN Verification': is_valid}
    
    df1 = pd.DataFrame(data)
    df = pd.merge(df, df1, how = 'left', on='IBAN',suffixes=('', '_y'))
    
    df.drop_duplicates(subset=['IBAN'], keep='first', inplace = True)
    
    
    st.write("------------------------------------------------------------------------------------")            
    
    st.write("Step 3: Download Extracted Bank Data")

    
    # buffer to use for excel writer
    buffer = io.BytesIO()
    
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    # Write each dataframe to a different worksheet.
        df.to_excel(writer, sheet_name='Sheet1', index = False)
        # Close the Pandas Excel writer and output the Excel file to the buffer
        writer.close()
    
        download2 = st.download_button(
            label="Download data as Excel",
            data=buffer,
            file_name='IBAN_Output.xlsx',
            mime='application/vnd.ms-excel'
        )
        
