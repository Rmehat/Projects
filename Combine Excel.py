'''
Consolidate several excel files
load into dataframe
Output to a single excel file
W:/High Content Harness ECNs/ECN Read files
'''

import os
import pandas as pd
import glob

#Location of folder
data_file_folder = "//na.paccar.com/ppdren/PublicFiles/High Content Harness ECNs/ECN Read files"

#Grabbing all .xlsx (Excel) files from Folder
file_list = glob.glob("//na.paccar.com/ppdren/PublicFiles/High Content Harness ECNs/ECN Read files/*.xlsx")

#Changing Columns that were causing Issue by using Dictionary - {'Key':'Value'}
names = {'New Side Part':'New Side', 'Old Side Part': 'Old Side', '(Old Side) Harness 2.1M/1.9M/Both':'(Old Side) Harness 2.1m or other model',
'(New Side) Specific Service Variations':'(New Side) Specific Service Variations', 'ECN Firm Date':'ECN Firm Date', 'Phase':'Phase',
  '(New Side) Engineering Division':'(New Side) Engineering Division', 'ECN Number':'ECN Number', 'Harness Type':'Harness Type', 
  '(Old Side) Covered Division':'(Old Side) Covered Division', '(Old Side) Latest Build Date (MM/DD/YYYY)':'(Old Side) Latest Build Date'}

df = []
for file in file_list:
    print('Loading file {0}...'.format(file))
    data = pd.read_excel(os.path.join(data_file_folder, file), sheet_name = 'ECN Part Sets')
    data.rename(columns=names, inplace=True)
    df.append(data)

#List Number of Excel Sheets Loaded Successfully (Don't need to do this)
print(str(len(df)) + " files loaded")

#Merge all DataFrame Objects to a Single Table
df_master = pd.concat(df, axis=0)
df_master.to_excel(r'//na.paccar.com/ppdren/PublicFiles/High Content Harness ECNs/ECN Read files/Test1.xlsx', index=False)
