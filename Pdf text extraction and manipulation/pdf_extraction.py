#%%
import numpy as np
import pandas as pd
import PyPDF2
import os

print(os.getcwd())

#%% table 2.4
os.chdir("/Users/mano/Documents/WHOI_1/Data_Cleaning/files")

tab_2_4 = pd.read_excel("table_2.4.xlsx", sheet_name="Table_2.4")
#tab_2_4 = tab_2_4.drop(tab_2_4.index[0])
new_header = tab_2_4.iloc[0]
tab_2_4 = tab_2_4[1:]
tab_2_4.columns = new_header
#tab_2_4.iloc[:,1]

tab_1 =pd.read_excel("table_2.4.xlsx", sheet_name="Table_1")

final_tab_2_4 = pd.merge(tab_2_4, tab_1, left_on='Core', right_on='Core', how='left')

final_tab_2_4.to_excel('table_2.4_combined.xlsx')

#%% table 2.3

tab_2_3 = pd.read_excel("Table_2.3.xlsx", sheet_name="Sheet1")

final_tab_2_3 = pd.merge(tab_2_3, tab_1, left_on="Core", right_on="Core", how="left")

final_tab_2_3.columns = ['Core', 'Mg/Ca', 'Sr/Ca', 'Mg/Ca', 'Sr/Ca', 'Mg/Ca', 'Sr/Ca',
                         'Mg/Ca', 'Sr/Ca', 'Mg/Ca', 'Sr/Ca', 'Latitude', 'Longitude', 'Depth',
                         'Temperature', 'Salinity', 'CO32-', 'Estimated Seawater Cd',
                         'Estimated Seawater Zn', 'Conventional 14C age', 'NOSAMS Number']


final_tab_2_3.to_excel('table_2.3_combined.xlsx')

#%% table 3.2

tab_3_2 = pd.read_excel("table3.2.xlsx",sheet_name="Sheet1")

final_tab_3_2 = pd.merge(tab_3_2,tab_1, left_on="Core", right_on="Core", how="left")

final_tab_3_2.columns = ['Core', 'Cd/Ca', 'Zn/Ca', 'Cd/Ca', 'Zn/Ca', 'Cd/Ca', 'Zn/Ca', 'Cd/Ca', 'Zn/Ca', 'Cd/Ca', 'Zn/Ca', 'Latitude', 'Longitude', 'Depth', 'Temperature', 'Salinity', 'CO32-', 'Estimated Seawater Cd', 'Estimated Seawater Zn', 'Conventional 14C age', 'NOSAMS Number']

final_tab_3_2.to_excel('table_3.2_combined.xlsx')

#%% added two rows manually in tab 2_3 in excel to make 3 files equal

v2_tab2_3 = pd.read_excel("table_2.3_combined.xlsx", sheet_name="Sheet1")

v2_tab2_3.columns = ['Core', 'Mg/Ca', 'Sr/Ca', 'Mg/Ca', 'Sr/Ca', 'Mg/Ca', 'Sr/Ca', 'Mg/Ca', 'Sr/Ca', 'Mg/Ca', 'Sr/Ca', 'Latitude', 'Longitude','Depth', 'Temperature', 'Salinity', 'CO32-', 'Estimated Seawater Cd','Estimated Seawater Zn', 'Conventional 14C age', 'NOSAMS Number']


v2_tab2_4 = final_tab_2_4
v2_tab3_2 = final_tab_3_2

#%%  C_pachyderma

C_pachyderma_2_3 = v2_tab2_3.iloc[:,0:3]
C_pachyderma_2_4 = v2_tab2_4.iloc[:,1:3]
C_pachyderma_3_2 = v2_tab3_2.iloc[:,1:3]

C_pachyderma = pd.concat([C_pachyderma_2_3, C_pachyderma_2_4,C_pachyderma_3_2 ], axis=1)
C_pachyderma = pd.merge(C_pachyderma, tab_1, left_on='Core', right_on='Core', how='left')


#%% U. peregrina

U_peregrina_2_3 = v2_tab2_3.iloc[:,[0,3,4]]
#U_peregrina_2_3 = pd.concat([v2_tab2_3.iloc[:,0],v2_tab2_3.iloc[:,3:5]])
U_peregrina_2_4 = v2_tab2_4.iloc[:,3:5]
U_peregrina_3_2 = v2_tab3_2.iloc[:,3:5]

U_peregrina = pd.concat([U_peregrina_2_3, U_peregrina_2_4,U_peregrina_3_2 ], axis=1)
U_peregrina = pd.merge(U_peregrina, tab_1, left_on='Core', right_on='Core', how='left')

#%% P. ariminensis

P_ariminensis_2_3 = v2_tab2_3.iloc[:,[0,5,6]]
P_ariminensis_2_4 = v2_tab2_4.iloc[:,5:7]
P_ariminensis_3_2 = v2_tab3_2.iloc[:,5:7]

P_ariminensis = pd.concat([P_ariminensis_2_3, P_ariminensis_2_4,P_ariminensis_3_2 ], axis=1)
P_ariminensis = pd.merge(P_ariminensis, tab_1, left_on='Core', right_on='Core', how='left')

#%% P. foveolata

P_foveolata_2_3 = v2_tab2_3.iloc[:,[0,7,8]]
P_foveolata_2_4 = v2_tab2_4.iloc[:,7:9]
P_foveolata_3_2 = v2_tab3_2.iloc[:,7:9]

P_foveolata = pd.concat([P_foveolata_2_3, P_foveolata_2_4,P_foveolata_3_2 ], axis=1)
P_foveolata = pd.merge(P_foveolata, tab_1, left_on='Core', right_on='Core', how='left')


#%% H. elegans


H_elegans_2_3 = v2_tab2_3.iloc[:,[0,9,10]]
H_elegans_2_4 = v2_tab2_4.iloc[:,9:11]
H_elegans_3_2 = v2_tab3_2.iloc[:,9:11]

H_elegans = pd.concat([H_elegans_2_3, H_elegans_2_4,H_elegans_3_2 ], axis=1)
H_elegans = pd.merge(H_elegans, tab_1, left_on='Core', right_on='Core', how='left')




#%% write to excel in multiple sheets

from pandas import ExcelWriter

list_dfs = [C_pachyderma, U_peregrina,P_ariminensis,P_foveolata,H_elegans]
xls_path = "/Users/mano/Documents/WHOI_1/Data_Cleaning/files"

def save_xls(list_dfs, xls_path):
    with ExcelWriter(xls_path) as writer:
        for n, df in enumerate(list_dfs):
            df.to_excel(writer,'sheet%s' % n)
        writer.save()

save_xls(list_dfs,xls_path)

#%%

C_pachyderma.to_excel("Combined_tables.xlsx")
U_peregrina.to_excel("Combined_tables", sheet_name='Sheet2')
P_ariminensis.to_excel("Combined_tables", sheet_name='Sheet3')
P_foveolata.to_excel("Combined_tables", sheet_name='Sheet4')
H_elegans.to_excel("Combined_tables", sheet_name='Sheet5')

import xlsxwriter

writer = pd.ExcelWriter('Combined_tables.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
C_pachyderma.to_excel(writer, sheet_name='C_pachyderma')
U_peregrina.to_excel(writer, sheet_name='U_peregrina')
P_ariminensis.to_excel(writer, sheet_name='P_ariminensis')
P_foveolata.to_excel(writer, sheet_name='P_foveolata')
H_elegans.to_excel(writer, sheet_name='H_elegans')

# Close the Pandas Excel writer and output the Excel file.
writer.save()
