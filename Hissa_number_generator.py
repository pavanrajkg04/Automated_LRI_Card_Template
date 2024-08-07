# -*- coding: utf-8 -*-
"""
Created on Thu May  7 15:24:55 2024

@author: Pavan Raj K G
"""

import pandas as pd
import pyodbc
import os


data= pd.read_excel(r"C:\Users\Admin\Desktop\reward_Raichur_cadastral.xlsx")
data.columns
df1 = pd.DataFrame(data[['village Code','MWS_ Code','SUJALA_LRI','surveynu_1']].astype(str))
#df1 = pd.DataFrame(data[['village Code','MWS_ Code','SUJALA_LRI']].astype(str))
df1['village Code'] = df1['village Code'].str[:10]
df1

for idx, vil in enumerate(df1['village Code']):
    if len(vil) <= 9:
       df1.loc[idx, 'village Code'] = '0' + vil
       
       
# Group data by 'MWS_ Code' and create lists of village codes
grouped_data = df1.groupby(['SUJALA_LRI','MWS_ Code','village Code'])['surveynu_1'].apply(list).reset_index()
#grouped_data = df1.groupby('MWS_ Code')['village Code'].apply(list).reset_index()
# Remove square brackets from each list of village codes
#grouped_data['surveynu_1'] = grouped_data['surveynu_1'].apply(lambda x: [code.replace('[', '').replace(']', '') for code in x])

# Define the connection parameters
server = '10.96.201.136'
database1 = 'KWDD_GIS'
database2 = 'KWDD_MIS'
username = 'mis'
password = 'sys@123'

# Create connection strings
connection_string1 = f'DRIVER=ODBC Driver 17 for SQL Server;SERVER={server};DATABASE={database1};UID={username};PWD={password}'
connection_string2 = f'DRIVER=ODBC Driver 17 for SQL Server;SERVER={server};DATABASE={database2};UID={username};PWD={password}'

data= pd.read_excel(r"C:\Users\Admin\Desktop\lriData.xlsx")


for index, row in grouped_data.iterrows():
    Partner = row['SUJALA_LRI']
    mws_code = row['MWS_ Code']
    village_codes = row['village Code']
    if isinstance(row['surveynu_1'], list):
        # Join the list elements to form a single string
        survey_string = ','.join(row['surveynu_1'])
        # Now you can perform the replace operation on the string
        survey_string = survey_string.replace("[", "").replace(']', '')
        # Optionally, you can assign the modified string back to the list if needed
        row['surveynu_1'] = survey_string
       
    survey_numbers = row['surveynu_1']
   
    # Process data from the current row
    print("Partner:", Partner)
    print("MWS Code:", mws_code)
    print("Village Codes:", village_codes)
    print("Survey numbers:", survey_numbers)
    query = f'''
 SELECT b.BhoomiDistrictName,b.BhoomiTalukName,b.BhoomiHobliName,b.BhoomiVillageName,
a.survey_no,
TRIM(a.surnoc) as surnoc,
TRIM(a.hissa_no) as hissa_no,
--TRIM(REPLACE(a.hissa_no,'*',' ')) as HISSA,
--dbo.RemoveSpecialChars(a.hissa_no) as hissa_new,
--TRIM(REPLACE(  trim(REPLACE(  trim(REPLACE(  trim(REPLACE(  trim(REPLACE( trim(REPLACE(a.hissa_no,'*',' ')),'\',' ')),'/',' ')),'+',' ')),'-',' ')),'.',' ')) as HISSA1,
--TRIM(a.hissa_no) as HISSA2,
trim(CONCAT(a.survey_no, '_', trim(REPLACE(a.hissa_no,'*',' ')))) as SurveyNo_Hissa,
   trim(CONCAT(a.survey_no, '_',a.owner_no,'_',a.survey_no, '_',   trim(REPLACE(  trim(REPLACE(  trim(REPLACE(  trim(REPLACE(  trim(REPLACE( trim(REPLACE(a.hissa_no,'*',' ')),'\',' ')),'/',' ')),'+',' ')),'-',' ')),'.',' '))  )) as  DXF_TEXT_old,
   trim(CONCAT(a.survey_no, '_',a.owner_no,'_',a.survey_no, '_',   trim(a.hissa_no)  )) as  DXF_TEXT,
a.owner_no,
TRIM(a.owner) as Farmer_Name,
TRIM(REPLACE(a.owner,'.',' ')) as Farmer_Name_1,
REPLACE(
  REPLACE(
   REPLACE(
    REPLACE(
REPLACE(
 REPLACE(
  REPLACE(
  REPLACE(
 REPLACE(
  REPLACE(
    REPLACE(
    REPLACE(
REPLACE(
REPLACE(
REPLACE(
REPLACE(
REPLACE(
REPLACE(
REPLACE(
REPLACE(
REPLACE(
REPLACE(
REPLACE(
 REPLACE(
  REPLACE(
           REPLACE(
               REPLACE(
 REPLACE(
      REPLACE(a.owner, 'M', N' ಎಂ '),
 'U', N' ಯು '),
      'I', N' ಐ '),
'V', N' ವಿ ' ),
'W', N' ಡಬ್ಲ್ಯೂ '),
'X', N' ಎಕ್ಸ್ '),
'Y', N' ವೈ '),
    'P', N' ಪಿ '),
'D', N' ಡಿ '),
'T', N' ಟಿ' ),
'R', N' ಆರ್ '),
'G', N' ಜಿ '),
'H', N' ಎಚ್ '),
'J', N' ಜೆ '),
 'L', N' ಎಲ್ '),
  'Smt', N' ಶ್ರೀಮತಿ '),
    'Sri', N' ಶ್ರೀ '),
   'E', N' ಇ '),
  'C', N' ಸಿ '),
  'O', N' ಓ '),
  'N', N' ಎನ್ '),
  'K', N' ಕೆ '),
  'A', N' ಎ '),
  'S', N' ಎಸ್ '),
   'Z', N' ಝಡ್‌ '),
   'B', N' ಬಿ ' ),
     'A', N' ಎ '),
   'F', N' ಎಫ್ '),
'.', ' ') as Farmer_Name_New,
TRIM(a.owner_sex) as owner_sex,
 CASE
    WHEN a.owner_sex='M' THEN N' ಗಂಡು '
    WHEN a.owner_sex='F' THEN N' ಹೆಣ್ಣು '
    WHEN a.owner_sex='O' THEN N' ಇತರೆ '  
END AS Gender,
TRIM(REPLACE(a.relationship,N'್','')) as relationship,
TRIM(REPLACE(a.father,N'್','')) as father,
trim(CONCAT(TRIM(a.relationship), ' ', TRIM(REPLACE(a.father,N'್','')))) as relationship_Father,
TRIM(b.KGISVillageCode) as KGISVillageCode
  FROM [KWDD_GIS].[dbo].[RTC_DATA] as a
  inner join [KWDD_MIS].[dbo].[KGIS_Bhoomi_Village_Master] as b on a.census_dist_code=b.BhoomiDistrictCode and a.census_taluk_code=b.[BhoomiTalukCode] and a.hobli_code=b.[BhoomiHobliCode] and a.village_code=b.VillageCode
   where a.owner_cat='PRV' and a.owner_no=a.main_owner_no and (a.[ext_acre]>0 OR a.[ext_gunta] >0 OR [ext_fgunta]>0) and b.KGISVillageCode in('{village_codes}') and a.survey_no in({survey_numbers});

'''

    try:
        # Connect to the database
        connection1 = pyodbc.connect(connection_string1)
       
        # Execute query and fetch data
        data = pd.read_sql(query, connection1)
        df = pd.DataFrame(data)
       
        # Export data to Excel
        df.to_excel(f'{Partner+" "+mws_code+" "+" "+village_codes}.xlsx', index=False)
        # Close the connection
        connection1.close()
   
    except Exception as e:
        print("An error occurred:", str(e))