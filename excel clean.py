# -*- coding: utf-8 -*-
"""
Created on Tue Nov 30 16:50:08 2021

@author: Ruby_Hu
"""

from openpyxl import load_workbook
import pandas as pd

df = pd.read_excel ('~\Desktop\\Ship Report_WK48_20211130.xlsx', sheet_name='ASGL SHIPPED')
df.insert(loc=2, column='Customer', value=['' for i in range(df.shape[0])])

CPM_list = ['MATIAS', 'ANDRES', 'BENZ', 'PABLO']
country_customer = {
    'AR': ['NEWSAN', 'COMPUMUN', 'MEGATONE', 'AIR', 'FRAVEGA', 'GARBARI', 'INFLUENCER', 'PCARTS', 'PILISAR'],
    'CL': ['FALABELLA', 'NEXSYS', 'IM', 'PCFACTORY', 'INTCOMEX', 'IDC']
    }


for key in country_customer:
    if country_customer[key] in df.loc[df['Revenue Country'] == key, 'Ship To']:
        