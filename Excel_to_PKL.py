import pickle as pk
import calendar
import pandas as pd
import numpy as np

cd = '/home/meltonzheng/Scripts/Excel_to_Google_Sheets/'

_df1 = pd.read_excel(r"/home/meltonzheng/Dropbox/Afya Dashboard/MCEI Tableau Data.xlsx", engine='openpyxl', header=None)
df1 = _df1.where(pd.notnull(_df1), None)

months = []
for L in df1[10][1:]:
    try:
        month = L[:str(L).find('/')]
        months.append(calendar.month_name[int(month)])
    except:
        months.append('No Month')

s_df1 = df1.drop(columns=[2, 3, 4, 5, 6, 7, 9, 13, 15, 16, 17, 18, 19, 20, 21, 22, 24, 25, 26, 28, 29, 30, 31, 32, 33, 34, 35, 37, 38, 39, 40, 41, 42, 43, 44, 45, 47, 48, 49, 50, 51])

list1 = s_df1.to_numpy().tolist()
list1[0].append('Month')
[list1[i+1].append(m) for i,m in enumerate(months)]

rngData1 = tuple(tuple(obj) for obj in list1)
with open(cd+'1.pkl', 'wb') as pickle_file:
    pk.dump(rngData1,pickle_file)

_df2 = pd.read_excel(r"/home/meltonzheng/Dropbox/Afya Dashboard/Source for Afya Earnings Dashboard.xlsx", engine='openpyxl', header=None)
df2 = _df2.where(pd.notnull(_df2), None)

months = []
for L in df2[3][1:]:
    try:
        month = L[:str(L).find('/')]
        months.append(calendar.month_name[int(month)])
    except:
        months.append('No Month')

list2 = df2.to_numpy().tolist()

list2[0].append('Month')
[list2[i+1].append(m) for i,m in enumerate(months)]

rngData2 = tuple(tuple(obj) for obj in list2)
with open(cd+'2.pkl', 'wb') as pickle_file:
    pk.dump(rngData2,pickle_file)

_df3 = pd.read_excel(r"/home/meltonzheng/Dropbox/eConsult Support and Follow-Up/MCEI and CDCR eConsult Follow-Up MASTER.xlsx", engine='openpyxl', header=None)
df3 = _df3.where(pd.notnull(_df3), None)

months = []
for L in df3[2][1:]:
    try:
        month = L[:str(L).find('/')]
        months.append(calendar.month_name[int(month)])
    except:
        months.append('No Month')

list3 = df3.to_numpy().tolist()

list3[0].append('Month')
[list3[i+1].append(m) for i,m in enumerate(months)]

rngData3 = tuple(tuple(obj) for obj in list3)
with open(cd+'3.pkl', 'wb') as pickle_file:
    pk.dump(rngData3,pickle_file)
