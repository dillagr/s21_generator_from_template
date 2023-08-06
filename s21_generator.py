#!/usr/bin/env python
# coding: utf-8

# In[1]:


## pip
from fillpdf import fillpdfs


# In[2]:


## regular conda libs
# import pandas as pd
# import numpy as np

from loguru import logger


# In[3]:


## built-in
import os

from pathlib import Path
from datetime import datetime


# In[4]:


## custom
from s21_data_from_spreadsheets import isolate_field_values, isolate_header_values, header_data_fr_spreadsheet


# In[5]:


## based on data from spreadsheet
df = header_data_fr_spreadsheet()
names = list(df['REPORT_NAME'].unique())


# In[6]:


## service year starts at September
def service_year() -> int:
    this = datetime.now()
    # logger.info(this)
    if this.month >= 9: return this.year +1
    else: return this.year


# In[8]:


## loop through list
for name in names:

    ## initialize the data dictionary with fields for update
    data_dict = fillpdfs.get_form_fields('S-21_E.pdf')
    # len(data_dict.keys())
    data_dict = { '900_11_Year_C': service_year() }

    ## body (or fields)
    fields_dict = isolate_field_values([name])
    data_dict.update(fields_dict)

    ## header
    header_dict = isolate_header_values([name])
    data_dict.update(header_dict)

    ## definitions
    outfile = data_dict.pop('outfile')
    input_pdf_path = os.path.join(Path().absolute(), "S-21_E.pdf")
    output_pdf_path = os.path.join(Path().absolute(), "exported", f"SY{service_year()}-{outfile}")

    ## output
    logger.debug(f"Generating: {output_pdf_path}")
    fillpdfs.write_fillable_pdf(input_pdf_path, output_pdf_path, data_dict, flatten=False)

    # break


# In[ ]:




