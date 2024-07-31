import numpy as np
import pandas as pd

# Replace 'your_file.xlsx' with the path to your Excel file
excel_file = pd.ExcelFile('Power BI inputs SQL.xlsm')

# Load a specific sheet into a DataFrame
source = excel_file.parse('Inputs')  # Replace 'Sheet1' with your sheet name or index
source = source.dropna(subset=['Sample_ID'])

# Now you can work with the data in 'df'
#print(df.head())  # Display the first few rows of the DataFrame

Boring_Info = source[['Boring_ID','Boring','Latitude','Longitude','Ground_Surface_Elevation']].drop_duplicates()
Boring_Info = Boring_Info.reset_index(drop=True)
Boring_Info = Boring_Info.dropna()
#Boring_Info

Soil_USCS = source[['Sample_ID','Soil_Description','USCS']]
Soil_USCS = Soil_USCS.reset_index(drop=True)
#Soil_USCS

Sample_Info = source[['Boring_ID','Sample_ID','Top_Elevation','Bottom_Elevation','Sample_Type']]
Sample_Info = Sample_Info.reset_index(drop=True)
#Sample_Info

df = source[['Sample_ID','Moisture_Content']]
MC_Info = df.dropna(subset=['Moisture_Content'])
MC_Info = MC_Info.reset_index(drop=True)
#MC_Info

df = source[['Sample_ID','N_Value']]
SPT_Info = df.dropna(subset=['N_Value'])
SPT_Info = SPT_Info.reset_index(drop=True)
#SPT_Info

df = source[['Sample_ID','PL','LL','PI']]
Atterburg_Info = df.dropna(subset=['PI'])
Atterburg_Info = Atterburg_Info.reset_index(drop=True)
#Atterburg_Info

df = source[['Sample_ID','Passing_200']]
Passing_200 = df.dropna(subset=['Passing_200'])
Passing_200 = Passing_200.reset_index(drop=True)
#Passing_200

df = source[['Sample_ID','General_Description','Lab_Verified']]
Sample_General = df.dropna(subset=['General_Description'])
Sample_General = Sample_General.reset_index(drop=True)
#Sample_General

#relationship
df1 = source[['Sample_ID','UW_Source','Uwm_SPT']]
df1.rename(columns={'Uwm_SPT': 'Moist_Unit_Weight'}, inplace=True)

#tested
df2 = source[['Sample_ID','Unit_Weight_source','Tested_Unit_Weight_pcf']]
df2.rename(columns={'Tested_Unit_Weight_pcf': 'Moist_Unit_Weight','Unit_Weight_source': 'UW_Source'}, inplace=True)

frames = [df1, df2]

result_dataframe = pd.concat(frames)
Unit_Weight = result_dataframe.dropna(subset=['Moist_Unit_Weight'])
Unit_Weight = Unit_Weight.reset_index(drop=True)
#Unit_Weight

#relationship
df1 = source[['Sample_ID','phi_Source',"phi"]]

#UU tested
df2 = source[['Sample_ID','UU_Source','UU_phi']]
df2.rename(columns={'UU_Source': 'phi_Source','UU_phi': 'phi'}, inplace=True)

#CU tested
df3 = source[['Sample_ID','CU_Source','CU_phi_deg']]
df3.rename(columns={'CU_Source': 'phi_Source','CU_phi_deg': 'phi'}, inplace=True)

#DS tested
df4 = source[['Sample_ID','DS_Source','DS_phi']]
df4.rename(columns={'DS_Source': 'phi_Source','DS_phi': 'phi'}, inplace=True)

frames = [df1, df2, df3, df4]

result_dataframe = pd.concat(frames)
Tot_phi = result_dataframe.dropna(subset=['phi'])
Tot_phi = Tot_phi.reset_index(drop=True)
#Tot_phi

#relationship
df1 = source[['Sample_ID','c_source','c_psf']]

#PP tested
df2 = source[['Sample_ID','PP_Source','PP_Su_psf']]
df2.rename(columns={'PP_Source': 'c_source','PP_Su_psf': 'c_psf'}, inplace=True)

#UC tested
df3 = source[['Sample_ID','UC_Source','UC_Su_psf']]
df3.rename(columns={'UC_Source': 'c_source','UC_Su_psf': 'c_psf'}, inplace=True)

#UU tested
df4 = source[['Sample_ID','UU_Source','UU_c_psf']]
df4.rename(columns={'UU_Source': 'c_source','UU_c_psf': 'c_psf'}, inplace=True)

#CU tested
df5 = source[['Sample_ID','CU_Source','CU_c_psf']]
df5.rename(columns={'CU_Source': 'c_source','CU_c_psf': 'c_psf'}, inplace=True)

#DS tested
df6 = source[['Sample_ID','DS_Source',"DS_eff_c_psf"]]
df6.rename(columns={'DS_Source': 'c_source',"DS_eff_c_psf": 'c_psf'}, inplace=True)

frames = [df1, df2, df3, df4, df5, df6]

result_dataframe = pd.concat(frames)
Tot_c = result_dataframe.dropna(subset=['c_psf'])
Tot_c = Tot_c.reset_index(drop=True)
#Tot_c

#relationship
df1 = source[['Sample_ID','clay_phi_Source',"Clay_eff_phi"]]
df1.rename(columns={'clay_phi_Source': 'eff_phi_source',"Clay_eff_phi": 'Eff_phi'}, inplace=True)

#Cu Test
df2 = source[['Sample_ID','CU_Source',"CU_eff_phi_psf"]]
df2.rename(columns={'CU_Source': 'eff_phi_source',"CU_eff_phi_psf": 'Eff_phi'}, inplace=True)

frames = [df1, df2]

result_dataframe = pd.concat(frames)
Eff_phi = result_dataframe.dropna(subset=['Eff_phi'])
Eff_phi = Eff_phi.reset_index(drop=True)
#Eff_phi

#relationship
df1 = source[['Sample_ID','e_source','e_Gs_2.67']]
df1.rename(columns={'e_Gs_2.67': 'void_ratio'}, inplace=True)

#UC Test
df2 = source[['Sample_ID','UC_Source',"UC_eo"]]
df2.rename(columns={'UC_Source': 'e_source',"UC_eo": 'void_ratio'}, inplace=True)

#UU Test
df3 = source[['Sample_ID','UU_Source',"UU_eo"]]
df3.rename(columns={'UU_Source': 'e_source',"UU_eo": 'void_ratio'}, inplace=True)

#CU Test
df4 = source[['Sample_ID','CU_Source',"CU_eo"]]
df4.rename(columns={'CU_Source': 'e_source',"CU_eo": 'void_ratio'}, inplace=True)

#Consolidation Test
df5 = source[['Sample_ID','CONS_Source',"CONS_eo"]]
df5.rename(columns={'CONS_Source': 'e_source',"CONS_eo": 'void_ratio'}, inplace=True)

frames = [df1, df2, df3, df4, df5]

result_dataframe = pd.concat(frames)
Void_ratio = result_dataframe.dropna(subset=['void_ratio'])
Void_ratio = Void_ratio.reset_index(drop=True)
#Void_ratio

#relationship
df1 = source[['Sample_ID','Cc_Source','Cc', 'Cr_Source', 'Cr', 'calpha_source', 'calpha', 'Cv_Source', 'Cv', "pc_Source", "pc"]]

#Consolidation Test Cc
df2 = source[['Sample_ID','CONS_Source',"CONS_Cc"]]
df2.rename(columns={'CONS_Source': 'Cc_Source',"CONS_Cc": 'Cc'}, inplace=True)

#Consolidation Test Cr
df3 = source[['Sample_ID','CONS_Source',"CONS_Cr"]]
df3.rename(columns={'CONS_Source': 'Cr_Source',"CONS_Cr": 'Cr'}, inplace=True)

#Consolidation Test Cv
df4 = source[['Sample_ID','CONS_Source',"CONS_Cv"]]
df4.rename(columns={'CONS_Source': 'Cv_Source',"CONS_Cv": 'Cv'}, inplace=True)

#Consolidation Test p'c
df5 = source[['Sample_ID','CONS_Source',"CONS_pc_ksf"]]
df5.rename(columns={'CONS_Source': "pc_Source","CONS_pc_ksf": "pc"}, inplace=True)

cons_merge = df2.merge(df3)
cons_merge = cons_merge.merge(df4)
cons_merge = cons_merge.merge(df5)

frames = [df1, cons_merge]

result_dataframe = pd.concat(frames)
Cons_Info = result_dataframe.dropna(subset=['Cc'])
Cons_Info = Cons_Info.reset_index(drop=True)
#Cons_Info

#relationship
df1 = source[['Sample_ID','Sand_Perm_Source','Perm_cm_s']]
df1.rename(columns={'Sand_Perm_Source': 'Perm_Source'}, inplace=True)

#Perm Test
df2 = source[['Sample_ID','PERM_Source',"PERM_cm_sec"]]
df2.rename(columns={'PERM_Source': 'Perm_Source',"PERM_cm_sec": 'Perm_cm_s'}, inplace=True)


frames = [df1, df2]

result_dataframe = pd.concat(frames)
Perm = result_dataframe.dropna(subset=['Perm_cm_s'])
Perm = Perm.reset_index(drop=True)
#Perm

#relationship
df1 = source[['Sample_ID','Es_Source','Es']]

Es = df1.dropna(subset=['Es'])
Es = Es.reset_index(drop=True)
#Es

import sqlite3

conn = sqlite3.connect('Geo_database.db')
cur = conn.cursor()

# Drop existing tables if they exist
tables = [
    'Boring_Info', 'Sample_Info', 'Soil_USCS', 'MC_Info', 
    'SPT_Info', 'Atterburg_Info', 'Passing_200', 'Sample_General', 
    'Unit_Weight', 'Tot_phi', 'Tot_c', 'Eff_phi', 'Void_ratio', 
    'Cons_Info', 'Perm', 'Es'
]

for table in tables:
    cur.execute(f'DROP TABLE IF EXISTS {table}')

# Enable foreign key support
cur.execute('PRAGMA foreign_keys = ON;')

cur.execute('''
CREATE TABLE Boring_Info (
    Boring_ID TEXT PRIMARY KEY,
    Boring TEXT,
    Latitude REAL,
    Longitude REAL,
    Ground_Surface_Elevation REAL                  
)         
''')

cur.execute('''
CREATE TABLE Sample_Info (
    Boring_ID TEXT,
    Sample_ID TEXT PRIMARY KEY,
    Top_Elevation REAL,
    Bottom_Elevation REAL,
    Sample_Type TEXT,
    FOREIGN KEY(Boring_ID) REFERENCES Boring_Info(Boring_ID)                 
)         
''')

cur.execute('''
CREATE TABLE Soil_USCS (
    Sample_ID TEXT,
    Soil_Description TEXT,
    USCS TEXT,
    FOREIGN KEY(Sample_ID) REFERENCES Sample_Info(Sample_ID)                 
)         
''')

cur.execute('''
CREATE TABLE MC_Info (
    Sample_ID TEXT,
    Moisture_Content REAL,
    FOREIGN KEY(Sample_ID) REFERENCES Sample_Info(Sample_ID)                  
)         
''')

cur.execute('''
CREATE TABLE SPT_Info (
    Sample_ID TEXT,
    N_Value INTEGER,
    FOREIGN KEY(Sample_ID) REFERENCES Sample_Info(Sample_ID)                 
)         
''')

cur.execute('''
CREATE TABLE Atterburg_Info (
    Sample_ID TEXT,
    PL REAL,
    LL REAL,
    PI REAL,
    FOREIGN KEY(Sample_ID) REFERENCES Sample_Info(Sample_ID)                 
)         
''')

cur.execute('''
CREATE TABLE Passing_200 (
    Sample_ID TEXT,
    Passing_200 REAL,
    FOREIGN KEY(Sample_ID) REFERENCES Sample_Info(Sample_ID)                  
)         
''')

cur.execute('''
CREATE TABLE Sample_General (
    Sample_ID TEXT,
    General_Description TEXT,
    Lab_Verified TEXT,
    FOREIGN KEY(Sample_ID) REFERENCES Sample_Info(Sample_ID)                
)         
''')

cur.execute('''
CREATE TABLE Unit_Weight (
    Sample_ID TEXT,
    UW_Source TEXT,
    Moist_Unit_Weight TEXT,
    FOREIGN KEY(Sample_ID) REFERENCES Sample_Info(Sample_ID)                  
)         
''')

cur.execute('''
CREATE TABLE Tot_phi (
    Sample_ID TEXT,
    phi_Source TEXT,
    phi REAL,
    FOREIGN KEY(Sample_ID) REFERENCES Sample_Info(Sample_ID)                  
)         
''')

cur.execute('''
CREATE TABLE Tot_c (
    Sample_ID TEXT,
    c_Source TEXT,
    c_psf REAL,
    FOREIGN KEY(Sample_ID) REFERENCES Sample_Info(Sample_ID)                  
)         
''')

cur.execute('''
CREATE TABLE Eff_phi (
    Sample_ID TEXT,
    eff_phi_source TEXT,
    Eff_phi REAL,
    FOREIGN KEY(Sample_ID) REFERENCES Sample_Info(Sample_ID)                  
)         
''')

cur.execute('''
CREATE TABLE Void_ratio (
    Sample_ID TEXT,
    e_source TEXT,
    void_ratio REAL,
    FOREIGN KEY(Sample_ID) REFERENCES Sample_Info(Sample_ID)                  
)         
''')

cur.execute('''
CREATE TABLE Cons_Info (
    Sample_ID TEXT,
    Cc_Source TEXT,
    Cc REAL,
    Cr_Source TEXT,
    Cr REAL,
    calpha_source TEXT,
    calpha REAL,
    Cv_Source TEXT,
    Cv REAL,
    pc_Source TEXT,
    pc REAL,
    FOREIGN KEY(Sample_ID) REFERENCES Sample_Info(Sample_ID)                  
)         
''')

cur.execute('''
CREATE TABLE Perm (
    Sample_ID TEXT,
    Perm_Source TEXT,
    Perm_cm_s REAL,
    FOREIGN KEY(Sample_ID) REFERENCES Sample_Info(Sample_ID)                  
)         
''')

cur.execute('''
CREATE TABLE Es (
    Sample_ID TEXT,
    Es_Source TEXT,
    Es REAL,
    FOREIGN KEY(Sample_ID) REFERENCES Sample_Info(Sample_ID)                  
)         
''')

# Clear tables
cur.execute('DELETE FROM Boring_Info')
cur.execute('DELETE FROM Soil_USCS')
cur.execute('DELETE FROM Sample_Info')
cur.execute('DELETE FROM MC_Info')
cur.execute('DELETE FROM SPT_Info')
cur.execute('DELETE FROM Atterburg_Info')
cur.execute('DELETE FROM Passing_200')
cur.execute('DELETE FROM Sample_General')
cur.execute('DELETE FROM Unit_Weight')
cur.execute('DELETE FROM Tot_phi')
cur.execute('DELETE FROM Tot_c')
cur.execute('DELETE FROM Eff_phi')
cur.execute('DELETE FROM Void_ratio')
cur.execute('DELETE FROM Cons_Info')
cur.execute('DELETE FROM Perm')
cur.execute('DELETE FROM Es')

# add data to created tables
Boring_Info.to_sql('Boring_Info', conn, if_exists='append', index=False)
Sample_Info.to_sql('Sample_Info', conn, if_exists='append', index=False)
Soil_USCS.to_sql('Soil_USCS', conn, if_exists='append', index=False)
MC_Info.to_sql('MC_Info', conn, if_exists='append', index=False)
SPT_Info.to_sql('SPT_Info', conn, if_exists='append', index=False)
Atterburg_Info.to_sql('Atterburg_Info', conn, if_exists='append', index=False)
Passing_200.to_sql('Passing_200', conn, if_exists='append', index=False)
Sample_General.to_sql('Sample_General', conn, if_exists='append', index=False)
Unit_Weight.to_sql('Unit_Weight', conn, if_exists='append', index=False)
Tot_phi.to_sql('Tot_phi', conn, if_exists='append', index=False)
Tot_c.to_sql('Tot_c', conn, if_exists='append', index=False)
Eff_phi.to_sql('Eff_phi', conn, if_exists='append', index=False)
Void_ratio.to_sql('Void_ratio', conn, if_exists='append', index=False)
Cons_Info.to_sql('Cons_Info', conn, if_exists='append', index=False)
Perm.to_sql('Perm', conn, if_exists='append', index=False)
Es.to_sql('Es', conn, if_exists='append', index=False)

# Commit and close the connection
conn.commit()
conn.close()