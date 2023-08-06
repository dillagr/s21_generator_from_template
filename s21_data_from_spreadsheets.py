#!/usr/bin/env python
# coding: utf-8

# In[1]:


## regular conda libs
import pandas as pd
import numpy as np

from loguru import logger
from dotenv import dotenv_values


# In[2]:


## built-in
import json
import os

from pathlib import Path
from datetime import datetime


# In[3]:


cfg = { **dotenv_values() }

## path where spreadsheets are located
PREFIX = os.path.join("C:\\", "Users", "great.dilla", "Google", "_____ KH-FILES _____", "_reports")


# In[4]:


## retrieve data for fields
def fields_data_fr_spreadsheet() -> pd.DataFrame:

    raw_df = pd.read_excel( os.path.join(PREFIX, cfg.get('fields_excel')), 
        sheet_name=cfg.get('fields_sheet'))

    column_names = {
        'Placements':'901',
        'Video showing':'902',
        'Hours':'903',
        'R.V.':'904',
        'BiStd':'905',
        }
    
    raw_df.rename(columns=column_names, inplace=True)
    # _ColsToDel = list(set(inactive.columns) - set(_ColsToRetain))

    return raw_df


# In[5]:


## retrieve data for headers
def header_data_fr_spreadsheet() -> pd.DataFrame:

    raw_df = pd.read_excel( os.path.join(PREFIX, cfg.get('headers_excel')), 
        sheet_name=cfg.get('headers_sheet'))

    column_names = {
        'BOX1':'900_4_CheckBox',
        'BOX2':'900_5_CheckBox',
        'BOX3':'900_6_CheckBox',
        'BOX4':'900_7_CheckBox',
        'BOX5':'900_8_CheckBox',
        'BOX6':'900_9_CheckBox',
        'BOX7':'900_10_CheckBox',
        }
    
    raw_df.rename(columns=column_names, inplace=True)
    # _ColsToDel = list(set(inactive.columns) - set(_ColsToRetain))

    return raw_df


# In[6]:


def isolate_header_values(publisher_identifier: list) -> dict:
    ## read the spreadsheet to pandas dataframe
    df = header_data_fr_spreadsheet()
    ## unpivot dataset
    df = df[df['REPORT_NAME'].isin(publisher_identifier)]                       ## filter publisher's name
    df = df.dropna(axis=1)                               ## blank MIDDLE_NAME becomes NaN (Not a Number)
    assert len(df) > 0
    
    df_dict = df.to_dict(orient='records')[-1]
    LAST_NAME = df_dict.pop('LAST_NAME')
    GIVEN_NAME = df_dict.pop('GIVEN_NAME')
    if "MIDDLE_NAME" in df_dict.keys():
        MIDDLE_NAME = df_dict.pop('MIDDLE_NAME')
    else:
        MIDDLE_NAME = ""
    df_dict['900_1_Text'] = f"{LAST_NAME}, {GIVEN_NAME} {MIDDLE_NAME}".strip()
    try:
        BIRTHDATE = df_dict.pop('BIRTHDATE')
        df_dict['900_2_Text'] = f'''      {BIRTHDATE.to_pydatetime().strftime("%d-%b-%Y").upper()}'''
    except: pass
    try:
        BAPTISM = df_dict.pop('BAPTISM')
        if BAPTISM: df_dict['900_3_Text'] = f'''  {BAPTISM.to_pydatetime().strftime("%d-%b-%Y").upper()}'''
    except:
        pass
    _ = df_dict.pop('LAST_GIVEN')
    _ = df_dict.pop('REPORT_NAME')

    __TODAY = datetime.now()
    if __TODAY.month > 8: YEAR = __TODAY.year +1 
    else: YEAR = __TODAY.year

    df_dict['outfile'] = f"SY{YEAR}-{LAST_NAME}-{GIVEN_NAME}-{MIDDLE_NAME}-S-21_E.pdf".strip().replace(" ","-").replace("--","-")
    ## return the dictionary
    return df_dict

# In[ ]:


def isolate_field_values(publisher_identifier: list) -> dict:
    ## read the spreadsheet to pandas dataframe
    raw_df = fields_data_fr_spreadsheet()
    ## unpivot dataset
    df = pd.melt(raw_df, id_vars=['Date', 'REPORT_NAME'], value_vars=['901', '902', '903', '904', '905'],
             var_name='Attribute', value_name='Value')
    df = df[df['REPORT_NAME'].isin(publisher_identifier)]                       ## filter publisher's name
    df = df[df['Value'].notna()]                                              ## remove NAN
    assert df.notnull().all().all()                                           ## check if all columns have values (not null or NaN)
    df['MonthNo'] = (df['Date'] +pd.DateOffset(months=4)).dt.month            ## offset by 4, sept=1 and so on
    df['Field'] = df.apply(lambda row: f"{row['Attribute']}_{row['MonthNo']}_S21_Value", axis=1)

    ## return the dictionary
    return dict(zip(df['Field'], df['Value']))


