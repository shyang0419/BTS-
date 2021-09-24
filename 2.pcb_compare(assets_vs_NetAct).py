#!/usr/bin/env python
# coding: utf-8

# In[4]:


# 將財產分配 load 程式
import pandas as pd
import numpy as np
import time
from datetime import date
import datetime
import re
import os


def parse_btsid(value):
    pattern = 'R\d{4,6}|\d{4,7}[uUlLnNgG]'
    m = re.search(pattern, value)
    if m and m.group(0):
        return m.group(0)
    else:
        return '沒填寫'

def detect_star(value):
    pattern = '([A-Za-z]{4,6})\*2|[A-Za-z0-9,;\[\]+]{4,24}'  # 還必須考慮,;[+]+等
    m = re.search(pattern, str(value))
    if m and m.group(1):
        return m.group(1) + ' ' + m.group(1)
    elif m and m.group(0):
        return m.group(0) 
    
def find_earlier_HW_Link():
    list1=['','']
    path= './data_HW_Link'
    date_list = os.listdir(path)
    n=0
    now1= date.today() 
    for i in range(30):   
        date_name_4G = '4G HW_Link_{}E.xlsm'.format(now1).replace('-','')
        date_name_5G = '5G HW_Link_{}E.xlsm'.format(now1).replace('-','')
        if date_name_4G in date_list:
            if list1[0] =='':
                list1[0] = date_name_4G
                n=n+1
        if date_name_5G in date_list:
            if list1[1] =='':
                list1[1] = date_name_5G
                n=n+1
        if n==2:
            break
        else:
            now1 = now1 - datetime.timedelta(days=1)
    return(list1)

def assets_duplicate(data):
    total_len=len(data)
    del_index=[]
    for i in range(total_len):
        if len(data[i])>1:
            res = data[i]
            del_index.append(res)
            output=[res[i:i + 1] for i in range(0, len(res), 1)]
            for j in range(len(output)):
                data.append(output[j])     

    for i in del_index:
        data.remove(i)
    data.sort()  
    return data

def find_last_record():
    list1=['']
    path= './results(analysis)'
    date_list = os.listdir(path)
    now= date.today() 
    now1 = now - datetime.timedelta(days=1)
    for i in range(30):   
        antenna_df = 'bts_circuitboard_switch({}).xlsx'.format(now1)
        if antenna_df in date_list:
            if list1[0] =='':
                list1[0] = antenna_df
            break
        else:
            now1 = now1 - datetime.timedelta(days=1)
    return(list1)

def find_earlier_in_stock():
    list1=['']
    path= './data_in_stock'
    date_list = os.listdir(path)
    now= date.today() 
    for i in range(120):   
        pcb_df = '電路板庫存({}).xlsx'.format(now)
        if pcb_df in date_list:
            if list1[0] =='':
                list1[0] = pcb_df
            break
        else:
            now = now - datetime.timedelta(days=1)
    return(list1)


# --------讀取 assets 基地台資料庫 (assets_df) ----------- 
today = date.today() 
file1_name = 'bts_figure({}).xlsx'.format(today)
file1_path = "./results(analysis)/{}".format(file1_name)
assets_df = pd.read_excel(file1_path, sheet_name = '基地台設備')
# ---------寫入資料庫  ---------------------------------
today = date.today() 
writer =pd.ExcelWriter('./results(analysis)/bts_circuitboard_switch({}).xlsx'.format(today))   

#---------調整 assets_df 格式 -----------
# 將 3G「FRGY」 電路板放入 4G 5G 財產中
assets_FRGY_df = assets_df[(assets_df['財產名稱'].isin(['3G行動電話收發訊系統'])) & 
                             (assets_df['型式/號'].str.contains("FRGY"))].copy()
assets_FRGY_df['型式/號'] = assets_FRGY_df['型式/號'].str.replace("Flexi-RRH","").str.replace('(','').str.replace(')','')

# index1 = assets_FRGY_df.loc[assets_FRGY_df['編號'].str.contains("U")].index   
# assets_FRGY_df.loc[index1 ,'設備狀態'] = '備援/備用' 

assets_FRGY_to_spare_df  = assets_FRGY_df[~assets_FRGY_df['編號'].str.contains("L")] 
index1 = assets_FRGY_df.loc[~assets_FRGY_df['編號'].str.contains("L")].index   
assets_FRGY_df.drop(index1, inplace = True)


assets_df = assets_df[assets_df['財產名稱'].isin(['4G行動寬頻系統共構設備','4G行動寬頻基地台','5G基地台射頻模組','5G基地台基頻模組'])]
assets_df = pd.concat([assets_df, assets_FRGY_df])
assets_df = assets_df.reset_index(drop = True)

assets_inuse_df = assets_df[(assets_df['設備狀態']=='使用中') & (assets_df['數量']== 1)]

assets_spare_df = assets_df[(assets_df['設備狀態']=='備援/備用') & (assets_df['數量']== 1)]
assets_spare_df = pd.concat([assets_spare_df,assets_FRGY_to_spare_df]) 
assets_spare_df  = assets_spare_df.reset_index(drop = True)

assets_stop_df = assets_df[(assets_df['設備狀態']=='停用') & (assets_df['數量']== 1)]
assets_loc_df = assets_df[(assets_df['設備狀態']=='佔位置') & (assets_df['數量']== 1)]

assets_inuse_df = assets_inuse_df[['編號','基地台名稱','型式/號','財產編號','異動者']]
assets_inuse_df['型式/號'] = assets_inuse_df['型式/號'].str.replace("AirScale","").str.replace('(','').str.replace(')','').str.replace('_','')
assets_inuse_df = assets_inuse_df.sort_values(by=['編號'])
assets_inuse_df.rename(columns={'基地台名稱':'基地台名(assets)','型式/號':'型式/號(assets)'},inplace =True)
assets_inuse_df.sort_values(by = ['編號','基地台名(assets)','型式/號(assets)'],inplace = True)
assets_inuse_df.to_excel(writer,sheet_name ='assets(使用中)',index=False)

assets_spare_df.to_excel(writer,sheet_name ='assets(備用)',index=False)
assets_stop_df.to_excel(writer,sheet_name ='assets(停用)',index=False)
assets_loc_df.to_excel(writer,sheet_name ='assets(佔位)',index=False)

#---------調整 netact_df 格式 ------------

# =====讀取 NetAct 基地台資料庫 (netact_df)=========#
n = 0
all_files = find_earlier_HW_Link() # 使用 def 
for filename in all_files:
    begin_df = pd.read_excel('./data_HW_Link/'+filename,sheet_name = 'Main',header = 5 )
    list_colname = list(begin_df.head())
    begin_df.rename(columns = {list_colname[0]:'編號',list_colname[1]:'基地台名稱',list_colname[4]:'硬體元件',list_colname[5]:'Serial Number'},inplace = True)
    begin_df = begin_df[['編號','基地台名稱','硬體元件','Serial Number']]
    begin_df['編號'] = begin_df['編號'].astype(str)
    begin_df['Serial Number'] = begin_df['Serial Number'].astype(str)
    begin_df['L_column'] = 'L'
    begin_df['N_column'] = 'N'
    if n == 0 :
        begin_df['編號'] = begin_df[['編號','L_column']].apply(''.join, axis=1)
        netact_df = begin_df
        n = n+1
    else:
        begin_df['編號'] = begin_df[['編號','N_column']].apply(''.join, axis=1)
        netact_df = netact_df.append(begin_df)
        
values = {'硬體元件': 'NULL','Serial Number': 'NULL'}
netact_df.fillna(value=values,inplace = True)

netact_df = netact_df.drop(['L_column','N_column'],axis = 1)  
netact_df = netact_df.reset_index(drop=True)

netact_df = netact_df.drop(netact_df.loc[netact_df['硬體元件'].str.contains(pat ='-1')].index)
netact_df = netact_df.drop(netact_df.loc[netact_df['硬體元件'].str.contains(pat ='473764A.102')].index)

netact_df = netact_df.drop(netact_df.loc[~netact_df['Serial Number'].str.contains(pat = '[A-Z]',regex= True)].index)

netact_df = netact_df[~netact_df['硬體元件'].isin(['FR2EB','FR2HB','FWEA_FREA','FWHN_FRHN'])]
netact_df['硬體元件'] = netact_df['硬體元件'].str.replace('ASIB AirScale Common','ASIB').str.replace('(','').str.replace(')','')   
netact_df['硬體元件'] = netact_df['硬體元件'].str.replace('AZQG S4-90M-R1-V2','AZQG').str.replace(' RV4S4-65A-R6','').str.strip()
netact_df['硬體元件'] = netact_df['硬體元件'].str.replace('AZQI S4-90M-R1-V3','AZQI')

netact_df = netact_df.dropna(subset =['基地台名稱'])
netact_df = netact_df.reset_index(drop = True)
netact_df.rename(columns={'基地台名稱':'基地台名(NetAct)','硬體元件':'型式/號(NetAct)'},inplace =True)
netact_df.sort_values(by=['編號','型式/號(NetAct)'],inplace = True)
netact_df = netact_df.reset_index(drop = True)
netact_df.to_excel(writer,sheet_name ='NetAct(使用中)',index=False)   


#---------合併 (assets_df) (netact_df) 兩資料庫內容 ------------
full_df = pd.concat([netact_df, assets_inuse_df])
full_df = full_df.reset_index(drop = True)
full_df = full_df[['編號','基地台名(NetAct)','型式/號(NetAct)','型式/號(assets)','財產編號','異動者']]
full_df.sort_values(by = ['編號','型式/號(NetAct)','型式/號(assets)'],inplace = True)
full_df = full_df.reset_index(drop = True)


# --------比較 (assets_df) (netact_df)兩資料庫內容 ----------- 
full_df['型式/號(assets)'] = full_df['型式/號(assets)'].map(detect_star)
full_df['型式/號(assets)'] = full_df['型式/號(assets)'].str.replace(',',' ').str.replace(';',' ').str.replace('[','').str.replace(']','').str.replace('+',' ')
# full_df.to_excel(writer,sheet_name ='基地台設備調整')

bts_id = list(full_df['編號'].unique())
for i in bts_id:
    netact_tmp_df = full_df[(full_df['編號']== i)&(full_df['型式/號(assets)'].isnull())]
    check1 = sorted(list(netact_tmp_df['型式/號(NetAct)'].str.split(' ')))
    assets_tmp_df = full_df[(full_df['編號']== i)&(full_df['型式/號(NetAct)'].isnull())]
    check2 = sorted(list(assets_tmp_df['型式/號(assets)'].str.split(' ')))    
    check2 = assets_duplicate(check2)
    
    # --------去除掉 FBBA & FBBC 電路板的比較----
    check1 = [x for x in check1 if x !=['FBBA']]
    check1 = [x for x in check1 if x !=['FBBC']]
    check2 = [x for x in check2 if x !=['FBBA']]
    check2 = [x for x in check2 if x !=['FBBC']]    
    # ------------------------------------------ 
    
    if check1 != check2:
        index1 = full_df.loc[full_df['編號']==i].index   
        full_df.loc[index1 ,'check'] = 'X'

        del_a = []
        a = check2
        b = check1
        for j in a:
            if j in b:
                del_a.append(j)
                b.remove(j)
        for j in del_a:
            a.remove(j)           
    
        full_df.loc[index1 ,'缺料'] = str(b)
        full_df.loc[index1 ,'多餘'] = str(a)        
               
        
    else:
        index2 = full_df.loc[full_df['編號']==i].index   
        full_df.loc[index2 ,'check'] = 'O'
        
full_df.to_excel(writer,sheet_name ='設備調整',index=False)
worksheet = writer.sheets['設備調整']
worksheet.set_column("A:A",9)
worksheet.set_column("B:E",18)
# 成績值比較 先選以前值
last_record =find_last_record() # 使用 def Macro
last_record_df = pd.read_excel('./results(analysis)/'+ last_record[0],sheet_name = '完成數',dtype= {'日期': str})

grade_df = full_df.loc[full_df.check=='O'].copy()
grade_df['異動者'].fillna(method='bfill',inplace=True)
grade_df.drop(['基地台名(NetAct)','型式/號(NetAct)','型式/號(assets)','財產編號','check'],axis =1,inplace=True)
grade_df = grade_df.drop_duplicates()
grade1_df = pd.DataFrame(grade_df['異動者'].value_counts()).T
grade1_df['日期']=str(today)
grade1_df.insert(0, '日期', grade1_df.pop('日期'))

result = pd.concat([last_record_df, grade1_df],ignore_index= True)

result.tail(15).to_excel(writer,sheet_name ='完成數',index = False)
worksheet = writer.sheets['完成數']
worksheet.set_column("A:C",10)

#============建立統計表(begin)===================#
assets_inuse_count = assets_inuse_df['型式/號(assets)'].value_counts()
assets_inuse_stic = pd.DataFrame(assets_inuse_count)
assets_inuse_stic.rename(columns={'型式/號(assets)':'assets_使用中'},inplace=True)
assets_inuse_stic.index.name="電路板"

assets_spare_df['型式/號']= assets_spare_df['型式/號'].str.replace(';',',').str.replace('AirScale','').str.replace('(','').str.replace(')','').str.replace('/','')
assets_spare_count = assets_spare_df['型式/號'].value_counts()
assets_spare_stic = pd.DataFrame(assets_spare_count)
assets_spare_stic.rename(columns={'型式/號':'assets_備用'},inplace=True)
assets_spare_stic.index.name="電路板"

assets_spare_north =  assets_spare_df[assets_spare_df['使用單位']=='北嘉義基維股']
assets_spare_north_cnt = assets_spare_north['型式/號'].value_counts()
assets_spare_norstic = pd.DataFrame(assets_spare_north_cnt)
assets_spare_norstic.rename(columns={'型式/號':'北基備用'},inplace=True)
assets_spare_norstic.index.name="電路板"

assets_spare_south =  assets_spare_df[assets_spare_df['使用單位']=='南嘉義基維股']
assets_spare_south_cnt = assets_spare_south['型式/號'].value_counts()
assets_spare_soustic = pd.DataFrame(assets_spare_south_cnt)
assets_spare_soustic.rename(columns={'型式/號':'南基備用'},inplace=True)
assets_spare_soustic.index.name="電路板"

assets_stop1_df = assets_stop_df.copy()
assets_stop1_df['型式/號']= assets_stop1_df['型式/號'].str.replace('AirScale','').str.replace('(','').str.replace(')','')
assets_stop_count  = assets_stop1_df['型式/號'].value_counts()
assets_stop_stic = pd.DataFrame(assets_stop_count)
assets_stop_stic.rename(columns={'型式/號':'assets_停用'},inplace=True)
assets_stop_stic.index.name="電路板"

assets_stop_north =  assets_stop1_df[assets_stop1_df['使用單位']=='北嘉義基維股']
assets_stop_north_cnt = assets_stop_north['型式/號'].value_counts()
assets_stop_norstic = pd.DataFrame(assets_stop_north_cnt)
assets_stop_norstic.rename(columns={'型式/號':'北基停用'},inplace=True)
assets_stop_norstic.index.name="電路板"

assets_stop_south =  assets_stop1_df[assets_stop1_df['使用單位']=='南嘉義基維股']
assets_stop_south_cnt = assets_stop_south['型式/號'].value_counts()
assets_stop_soustic = pd.DataFrame(assets_stop_south_cnt)
assets_stop_soustic.rename(columns={'型式/號':'南基停用'},inplace=True)
assets_stop_soustic.index.name="電路板"
assets_stop_soustic

netact_df_count = netact_df['型式/號(NetAct)'].value_counts()
netact_df_stic = pd.DataFrame(netact_df_count)
netact_df_stic.rename(columns={'型式/號(NetAct)':'NetAct_使用中'},inplace=True)
netact_df_stic.index.name="電路板"

statistics_df = assets_inuse_stic.join(assets_spare_stic,how='outer')
statistics_df = statistics_df.join(assets_spare_norstic,how='outer')
statistics_df = statistics_df.join(assets_spare_soustic,how='outer')
statistics_df = statistics_df.join(assets_stop_stic,how='outer')
statistics_df = statistics_df.join(assets_stop_norstic,how='outer')
statistics_df = statistics_df.join(assets_stop_soustic,how='outer')
statistics_df = statistics_df.join(netact_df_stic,how='outer')
statistics_df = statistics_df.fillna(0)
statistics_df[['assets_使用中','assets_備用','北基備用','南基備用','assets_停用','北基停用','南基停用','NetAct_使用中']] = statistics_df[['assets_使用中','assets_備用','北基備用','南基備用','assets_停用','北基停用','南基停用','NetAct_使用中']].astype(int)


#============讀取電路板庫存量(in stock)================#
file1_name = find_earlier_in_stock() # 使用 def 
file1_path = "./data_in_stock/{}".format(file1_name[0])
in_stock_df = pd.read_excel(file1_path, sheet_name = '統計')
in_stock_df.fillna(value=0, inplace=True)
in_stock_df['庫存_嘉義']= in_stock_df['庫存_北基']+in_stock_df['庫存_南基']
in_stock_df.set_index("電路板" , inplace=True)

statistics_df = statistics_df.join(in_stock_df,how='outer')
statistics_df = statistics_df[['assets_使用中','assets_備用','北基備用','南基備用','assets_停用','北基停用','南基停用','NetAct_使用中','庫存_嘉義','庫存_北基','庫存_南基']]
statistics_df.fillna(value=0, inplace=True)
statistics_df['財編缺額'] = statistics_df['NetAct_使用中'] + statistics_df['庫存_嘉義'] - statistics_df['assets_使用中'] - statistics_df['assets_備用'] - statistics_df['assets_停用'] 

statistics_df = statistics_df.reset_index()
stat_style = statistics_df.style.applymap(lambda x: 'background-color:#ADD8E6', subset=["北基備用"])     .applymap(lambda x: 'background-color:#ADD8E6', subset=["庫存_北基"])     .applymap(lambda x: 'background-color:#FFFF74', subset=["assets_停用"])     .applymap(lambda x: 'background-color:#FFFF74', subset=["北基停用"])     .applymap(lambda x: 'background-color:#FFFF74', subset=["南基停用"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["南基備用"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_南基"]) 
stat_style.to_excel(writer,sheet_name ='統計表',index = False)
worksheet = writer.sheets['統計表']
worksheet.set_column("A:A",10)
worksheet.set_column("B:B",15)
worksheet.set_column("C:H",12)
worksheet.set_column("I:I",15)
worksheet.set_column("J:M",12)

#----------------統計表 大於 > 0  ---------------------#
# 1.assets_使用中 < Netact_使用中 ，assets_備用 &assets_停用 !=0  提出
# 2.nmoss_使用中 == Netact_使用中 
# 備用 < 庫存， 停用 !=0 提出
# 備用 > 庫存  提出
# 3.assets_使用中 > Netact_使用中 提出

A = statistics_df['財編缺額'] > 0
stat_great_0_df = statistics_df[A]

B = stat_great_0_df['assets_使用中'] < stat_great_0_df['NetAct_使用中']
C = stat_great_0_df['assets_備用'] != 0
D = stat_great_0_df['assets_停用'] != 0
E = B&(C|D)

F = stat_great_0_df['assets_使用中'] == stat_great_0_df['NetAct_使用中']
G = stat_great_0_df['南基備用'] < stat_great_0_df['庫存_南基']
H = stat_great_0_df['南基停用'] != 0 
I = stat_great_0_df['北基備用'] < stat_great_0_df['庫存_北基']
J = stat_great_0_df['北基停用'] != 0
K = stat_great_0_df['南基備用']> stat_great_0_df['庫存_南基']
L = stat_great_0_df['北基備用']> stat_great_0_df['庫存_北基']
M = F&((G&H)|(I&J)|K|L)

N = stat_great_0_df['assets_使用中'] > stat_great_0_df['NetAct_使用中']

stat_great_1_df = stat_great_0_df[E|M|N]

great_style = stat_great_1_df.style.applymap(lambda x: 'background-color:#ADD8E6', subset=["assets_備用"])       .applymap(lambda x: 'background-color:#ADD8E6', subset=["assets_停用"])       .applymap(lambda x: 'background-color:#FFFF74', subset=["assets_使用中"])       .applymap(lambda x: 'background-color:#FFFF74', subset=["NetAct_使用中"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["北基備用"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["南基備用"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_北基"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_南基"])

great_style.to_excel(writer,sheet_name ='缺額>0修正',index = False)
worksheet = writer.sheets['缺額>0修正']
worksheet.set_column("A:A",10)
worksheet.set_column("B:B",15)
worksheet.set_column("C:H",12)
worksheet.set_column("I:I",15)
worksheet.set_column("J:M",12)

#----------------統計表 小於 < = 0  ---------------------#
# 1.assets_使用中 < NetAct_使用中  提出
# 2.assets_使用中 == NetAct_使用中 
# 備用 < 庫存， 停用 !=0 提出
# 備用 > 庫存  提出
# 3.assets_使用中 > NetAct_使用中 提出

A = statistics_df['財編缺額'] <= 0

stat_less_0_df = statistics_df[A]

B = stat_less_0_df['assets_使用中'] < stat_less_0_df['NetAct_使用中']

C = stat_less_0_df['assets_使用中'] == stat_less_0_df['NetAct_使用中']
D = stat_less_0_df['南基備用'] < stat_less_0_df['庫存_南基']
E = stat_less_0_df['南基停用'] != 0 
F = stat_less_0_df['南基備用'] >  stat_less_0_df['庫存_南基']
G = stat_less_0_df['北基備用'] < stat_less_0_df['庫存_北基']
H = stat_less_0_df['北基停用'] != 0 
I = stat_less_0_df['北基備用'] >  stat_less_0_df['庫存_北基']
J = C&((D&E)|F|(G&H)|I)

K = stat_less_0_df['assets_使用中'] > stat_less_0_df['NetAct_使用中']

stat_less_1_df = stat_less_0_df[B|J|K]

less_style = stat_less_1_df.style.applymap(lambda x: 'background-color:#ADD8E6', subset=["assets_使用中"])     .applymap(lambda x: 'background-color:#ADD8E6', subset=["NetAct_使用中"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["北基備用"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["南基備用"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_北基"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_南基"]) 

less_style.to_excel(writer,sheet_name ='缺額<=0修正',index = False)
worksheet = writer.sheets['缺額<=0修正']
worksheet.set_column("A:A",10)
worksheet.set_column("B:B",15)
worksheet.set_column("C:H",12)
worksheet.set_column("I:I",15)
worksheet.set_column("J:M",12)

#===============建立統計表(end)========================#

# 關閉寫入檔案
writer.save()


# In[ ]:





# In[ ]:




