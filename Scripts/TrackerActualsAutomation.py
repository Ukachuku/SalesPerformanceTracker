# -*- coding: utf-8 -*-
"""
Created on Mon Apr 27 09:55:28 2020

@author: Julian.Haro
"""

import pyodbc
import pandas as pd
import os
import xlwings as xw
from xlwings.constants import DeleteShiftDirection
from xlwings import Range, constants
#os.chdir('Y:\\Sales Performance Tracker Automation\\extract')

os.chdir('T:\\Julian.Haro\\Sales Performance Tracker Automation\\extract')

wb = xw.Book('T:\\Julian.Haro\\Sales Performance Tracker Automation\\extract\\Detail by Sale Order.xlsx')#replace with today_variable when adding this to main
wb2 = xw.Book('T:\\Julian.Haro\\Sales Performance Tracker Automation\\Sales_Update_Template\\Sales_Update_Template.xlsx')
wb3 = xw.Book('T:\\Julian.Haro\\Sales Performance Tracker Automation\\Sales Performance Tracker_2020\\Sales Performance Tracker_2020.xlsx')
wb4 = xw.Book('T:\\Julian.Haro\\Sales Performance Tracker Automation\\ActualVsSalesWk\\Actual vs Sales.xlsx')

#Instantiate a sheet object:
sht = wb.sheets['Sheet1']
sht1 = wb2.sheets['DB_Shipped not Invoiced']
sht2 = wb2.sheets['Pivot_Sales']
sht3 = wb3.sheets['Conv Sales by Mkt']

lastrowofextract = wb2.sheets[0].range('J' + str(wb2.sheets[0].cells.last_cell.row)).end('up').row
wb4.sheets[0].range('A'+str(lastrowofforecast+1)+':I'+str(lastrowofactuals)).clear()

df = pd.read_excel('T:\\Julian.Haro\\Sales Performance Tracker Automation\\extract\\Detail by Sale Order.xlsx', header=2, index=False)
df = df.drop(columns=['saleprice', 'saleamt', 'Price Category (groups)', 'Grades', 'warehouse', 'FY', 'Holds'])

#REMOVE ORDERS
df = df[~df.Customer.str.contains("HOLD")]
df = df[~df.Customer.str.contains("CASH SALE")]
df = df[~df.Customer.str.contains("DUMPED INVENTORY")]
df = df[~df.Customer.str.contains("TRADESHOW")]
df = df[~df.Customer.str.contains("PROSPECT SAMPLES")]
df = df[~df.Customer.str.contains("EMPLOYEE GIVE AWAY")]
df = df[~df.Customer.str.contains("LOST INVENTORY")]
df = df[~df.Customer.str.contains("BAGGING")]
df = df[~df.Customer.str.contains("COND FORECAST")]

#SLICE SIZE STRING CONTAINING PERIODS
df['Size'] = df.Size.str.slice(start=0, stop=2)

#convert size 96 to PW
df['Size'] = df.Size.replace(to_replace='96', value='PW')

#SLICE GRADE STRING TO GET LEFT TWO CHARACTERS
df['Grade'] = df.Grade.replace(to_replace='#1 COND', value='#1')
df['Grade'] = df.Grade.replace(to_replace='#1 001', value='#1')
df['Grade'] = df.Grade.replace(to_replace='#1 002', value='#1')
df['Grade'] = df.Grade.replace(to_replace='#1 001COND', value='#1')
df['Grade'] = df.Grade.replace(to_replace='#1 002COND', value='#1')

df['Grade'] = df.Grade.replace(to_replace='#2 COND', value='#2')
df['Grade'] = df.Grade.replace(to_replace='#2 001', value='#2')
df['Grade'] = df.Grade.replace(to_replace='#2 002', value='#2')
df['Grade'] = df.Grade.replace(to_replace='#2 001COND', value='#2')
df['Grade'] = df.Grade.replace(to_replace='#2 002COND', value='#2')

#copy transformed extract into the template
sht1.range('A3').options(index=False).value = df

#refresh pivot
wb2.api.RefreshAll()

#######MANUAL copy range of MX & CA #1 & #2 AA81:AA88 to 'Sales Performance Tracker_2020.xlsx' MANUAL due to column update of everyweek

#ACTUALS VS SALES WEEK# LOGIC
df = pd.read_excel('T:\\Julian.Haro\\Sales Performance Tracker Automation\\extract\\Detail by Sale Order.xlsx', header=2, index=False)
vlookupdf = pd.read_excel('T:\\Julian.Haro\\Sales Performance Tracker Automation\\VlookupData\\data.xlsx', index=False)
vlookupdf1 = pd.read_excel('T:\\Julian.Haro\\Sales Performance Tracker Automation\\VlookupData\\data.xlsx', sheet_name='Sheet3', index=False)

#REMOVE ORDERS
df = df[~df.Customer.str.contains("HOLD")]
df = df[~df.Customer.str.contains("CASH SALE")]
df = df[~df.Customer.str.contains("DUMPED INVENTORY")]
df = df[~df.Customer.str.contains("TRADESHOW")]
df = df[~df.Customer.str.contains("PROSPECT SAMPLES")]
df = df[~df.Customer.str.contains("EMPLOYEE GIVE AWAY")]
df = df[~df.Customer.str.contains("LOST INVENTORY")]
df = df[~df.Customer.str.contains("BAGGING")]
df = df[~df.Customer.str.contains("COND FORECAST")]

#select columns
df1 = df[['COO','Sales Type', 'Grade', 'Customer', 'Size', 'invcdescr', 'Pallets', 'sono']]

#replace grade values
df1['Grade'] = df1.Grade.replace(to_replace='#1 COND', value='#1')
df1['Grade'] = df1.Grade.replace(to_replace='#1 001', value='#1')
df1['Grade'] = df1.Grade.replace(to_replace='#1 002', value='#1')
df1['Grade'] = df1.Grade.replace(to_replace='#1 001COND', value='#1')
df1['Grade'] = df1.Grade.replace(to_replace='#1 002COND', value='#1')

df1['Grade'] = df1.Grade.replace(to_replace='#2 COND', value='#2')
df1['Grade'] = df1.Grade.replace(to_replace='#2 001', value='#2')
df1['Grade'] = df1.Grade.replace(to_replace='#2 002', value='#2')
df1['Grade'] = df1.Grade.replace(to_replace='#2 001COND', value='#2')
df1['Grade'] = df1.Grade.replace(to_replace='#2 002COND', value='#2')

#SLICE SIZE STRING CONTAINING PERIODS   
df1['Size'] = df1.Size.str.slice(start=0, stop=2) 
   
#merge on product description and invcdescr
mergedf = df1.merge(vlookupdf, how='left', left_on='invcdescr', right_on='Product Description')

######MANUAL update the product list on data workbook in vlookupdata directory then run this code below to get join for every record
vlookupdf = pd.read_excel('T:\\Julian.Haro\\Sales Performance Tracker Automation\\VlookupData\\data.xlsx', index=False)
mergedf = df1.merge(vlookupdf, how='left', left_on='invcdescr', right_on='Product Description')
#######MANUAL#if all records are joined proceed

#filter out organics; conventionals only
x = mergedf[mergedf['Sales Type']=='Conventional']

#filter out #2 ripe and culls
y = x[(x['Grade_x']!='Culls') & (x['Grade_x']!='#2 Ripe') & (x['Grade_x']!='#1 Ripe') ]

#transforming the data further
y = y[['COO','Sales Type', 'Grade_x', 'Size_x', 'Style', 'Customer', 'Pallets', 'sono']]
z = y.merge(vlookupdf1, how='left', left_on='Style', right_on='STYLE')
#######MANUAL: update the vlookup table for missing styles
a = z[['COO', 'Sales Type', 'Grade_x', 'Size_x', 'Style', 'Customer', 'PACK TYPE', 'CATEGORY', 'Pallets']]
#Select columns to paste to actual vs sales wk workbook
lastrowofforecast = wb4.sheets[0].range('J' + str(wb4.sheets[0].cells.last_cell.row)).end('up').row
lastrowofactuals = wb4.sheets[0].range('I' + str(wb4.sheets[0].cells.last_cell.row)).end('up').row
wb4.sheets[0].range('A'+str(lastrowofforecast+1)+':I'+str(lastrowofactuals)).clear()
wb4.sheets[0].range('A'+str(lastrowofforecast+1)).options(pd.DataFrame, index=False, header=False).value = a
#######MANUAL: for empty values in G & H, perform a vlookup on 'DATA' tab to get the pack type and category (note: the 'Data' tab might be hidden)
#######MANUAL: go to inventory projections workbook, found at T:\Supply Chain\Supply Chain Projections use the most current workbook. normally its from the week number prior to the current week e.g if we are in week 19 then use 'Inventory Projections Week 18_2020_Product Detail.xlsx'
#######MANUAL: once in the workbook go to 'Demand_Summary' tab go to current week and copy the adjustment column value to the RIGHT SIZE, COMMODITY, AND TYPE
#######MANUAL: pivot
#######Manual: make sure pack type and category do not have error values, some of the common errors are RPCSI missing the 'S' or Lug 12-2 missing a 'LB'
#######Manual: lookout for any truncation in customer field
