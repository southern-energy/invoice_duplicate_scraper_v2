# Started by Gregory Power at 11/06/19 @ 4:27 PM
# # Basic Functionality Achieved on 11/12/19
# # Able to Print Indexes on 11/19/19
# # Able to Search for ServiceID or Address + ServiceType Duplicates on 12/11/19 at 4:42 PM
# For Environment, use Python 3.6.0
# Working Tree Change

# We need to find duplicates that exist, there is one in the master sheet at the bottom of the most recent sheet.


import xlrd
import xlwt
import xlutils # Module Requires both XLRD & XLWT to be imported.
import pandas as pd
import numpy as np

# =======================================================
# =======================================================

# The current_sheet variable needs to be named the sheet you want to check for Duplicate (Service Types &Addresses) OR duplicate (ServiceIDs).

current_sheet = 'C:/Users/SEM/Documents/Invoicing/2-25-21.xlsx'

large_sheet = 'S:/SEM/Building Science Team/Accounting/Invoicing Spreadsheets/2020 Inspections Billing.xlsx'

large_sheet_2 = 'S:/SEM/Building Science Team/Accounting/Invoicing Spreadsheets/2021 Inspections Billing.xlsx'

# =======================================================
# =======================================================
# =======================================================
# =======================================================
# =======================================================

# Find Master Sheet

#workbook = xlrd.open_workbook(large_sheet)
pd.read_excel(large_sheet)
pd.read_excel(large_sheet_2)

# Start of Address + Service Type Duplicate Section


# Find All of the Sheets in the Workbook
# Combine all sheets of Master Sheet into a single list of lists.

df_master_Street_Address_And_Service = pd.concat(pd.read_excel(large_sheet, sheet_name=None, usecols=[1, 10], skiprows=0,header=None), sort=False, ignore_index=False)

df_master_Street_Address_And_Service_2 = pd.concat(pd.read_excel(large_sheet_2, sheet_name=None, usecols=[1, 10], skiprows=0,header=None), sort=False, ignore_index=False)

# This should combine the two dataframes on top of each other and should then run as normal.

df_master_Street_Address_And_Service = pd.concat([df_master_Street_Address_And_Service,df_master_Street_Address_And_Service_2], axis=0)

# print(df_master_Street_Address_And_Service)

# Read all of the sheets, using just the columns that have Street Address and the Service.

df_current_sheet_Street_Address_And_Service = pd.concat(pd.read_excel(current_sheet, sheet_name=None, usecols=[1, 10], skiprows=0,header=None), sort=False, ignore_index=False)

# print(df_current_sheet_Street_Address_And_Service)

# First Have to make them into a list of lists (https://stackoverflow.com/questions/22341271/get-list-from-pandas-dataframe-column)

ser_aggRows = pd.Series(df_master_Street_Address_And_Service.values.tolist())


# print('Printing: ser_aggRows (This collapses each row in Excel into a row this script can read.)',ser_aggRows, sep='\n', end='\n\nWe have finished organizing the rows of the Master Workbook\'s sheets into lists.\n\n\n')

ser_aggRows_current_sheet = pd.Series(df_current_sheet_Street_Address_And_Service.values.tolist())

# print('Printing: ser_aggRows_current_sheet (This collapses each row in Excel into a row this script can read.)',ser_aggRows_current_sheet, sep='\n', end='\n\nWe have finished organizing the rows of the Current Workbook\'s sheets into lists.\n\n')

first_set = set(map(tuple, ser_aggRows))
secnd_set = set(map(tuple, ser_aggRows_current_sheet))
second_set_storage = (map(tuple, ser_aggRows_current_sheet))


duplicates = first_set.intersection(secnd_set)

if len(duplicates) > 0:
    print("\n\nAddress and Service Type duplicates are: \n==========================================================================", duplicates, sep="\n", end="\n==========================================================================\nPlease use the list of Address and Service type pair(s) above to find the Address + Service Type duplicates.\n==========================================================================\n")
else: 
    print("\n==========================================================================\nThere are no Address + Service Type duplicates!\n==========================================================================\n")

# Duplicates_list converts the duplicates (an object type: set, with tuples inside it, to a list of lists again)

duplicates_list = list(map(list, duplicates))

secnd_set_list = list(map(list, second_set_storage))

# print(duplicates_list)
# print("The Length of the Second Set List: " + str(len(secnd_set_list)))
# print("The Length of the Aggrows Sheet: " + str(len(ser_aggRows_current_sheet)))


k = 0
Excel_Indexes = []
while k < len(duplicates_list):
    # print(secnd_set_list.index(duplicates_list[k]))
    Excel_Indexes.append(int(1) + int(secnd_set_list.index(duplicates_list[k])))
    k += 1
else:
    Excel_Indexes.sort()
    print("\nWe are done. Look to the following rows in Excel for Address + Service Type duplicates: \n", Excel_Indexes, sep="\n")

# End of Address + Service Type Duplicate Section

# Start of Service ID Section

# Master Sheet

df_master_ServiceID = pd.concat(pd.read_excel(large_sheet, sheet_name=None, usecols=[7], skiprows=0,header=None), sort=False, ignore_index=False, join="outer")

df_master_ServiceID_2 = pd.concat(pd.read_excel(large_sheet_2, sheet_name=None, usecols=[7], skiprows=0,header=None), sort=False, ignore_index=False, join="outer")

df_master_ServiceID.fillna(0, inplace = True)

df_master_ServiceID_2.fillna(0, inplace = True)

df_master_Street_Address_And_Service = pd.concat([df_master_ServiceID,df_master_ServiceID_2], axis=0)

# Current Sheet

df_current_sheet_ServiceID = pd.concat(pd.read_excel(current_sheet, sheet_name=None, usecols=[7], skiprows=0,header=None), sort=False, ignore_index=False)

#df_master_serviceID = df_master_ServiceID.astype(int)

# Master Sheet

ser_aggRows_master_ServiceID = pd.Series(df_master_ServiceID.values.tolist())

# print(ser_aggRows_master_ServiceID)

#Current Sheet
ser_aggRows_current_sheet_ServiceID = pd.Series(df_current_sheet_ServiceID.values.tolist())

# print(ser_aggRows_current_sheet_ServiceID)

# Master Sheet
first_set_ServiceID = set(map(tuple, ser_aggRows_master_ServiceID))

#print(first_set_ServiceID)

secnd_set_ServiceID = set(map(tuple, ser_aggRows_current_sheet_ServiceID))

#print(secnd_set_ServiceID)

second_set_storage_ServiceID = (map(tuple, ser_aggRows_current_sheet_ServiceID))

duplicates_ServiceID = first_set_ServiceID.intersection(secnd_set_ServiceID)

duplicates_list_ServiceID_Duplicates = list(map(list, duplicates_ServiceID))

secnd_set_list_ServiceID_Duplicates = list(map(list, second_set_storage_ServiceID))

if len(duplicates_ServiceID) > 0:
    print("\nDuplicates for ServiceID Condition are: \n==========================================================================", duplicates_ServiceID, sep="\n", end="\n==========================================================================\nPlease use the list of ServiceIDs above to find the duplicate ServiceIDs.\n")
else: 
    print("\n==========================================================================\nThere are no ServiceID duplicates!\n==========================================================================\n")


q = 0
Excel_Indexes_for_ServiceID_Duplicates = []
while q < len(duplicates_list_ServiceID_Duplicates):
    # print(secnd_set_list_Address_Service_Duplicates.index(duplicates_list_ServiceID_Duplicates[q]))
    Excel_Indexes_for_ServiceID_Duplicates.append(int(1) + int(secnd_set_list_ServiceID_Duplicates.index(duplicates_list_ServiceID_Duplicates[q])))
    q += 1
else:
    Excel_Indexes_for_ServiceID_Duplicates.sort()
    print("\nWe are done. Look to the following rows in Excel for Service ID Duplicates: \n", Excel_Indexes_for_ServiceID_Duplicates)

# End of Service ID Section;
