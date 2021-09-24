#!/usr/bin/env python
# coding: utf-8

# In[4]:


# 將財產分配 load 程式
import pandas as pd
import numpy as np
import time
from datetime import date
import datetime
import os
import re


def parse_btsid(value):
    pattern = 'R\d{4,6}|\d{4,7}[uUlLnNgG]'
    m = re.search(pattern, value)
    if m and m.group(0):
        return m.group(0)
    else:
        return '沒填寫'

def choice_char(value):
    pattern = '([1-9])[0-9][0-9]' 
    m = re.search(pattern, str(value))
    if m and m.group(1):
        return m.group(1)
    
def find_earlier_eNodebList():
    list1=['','']
    path= './data_eNodebList'
    date_list = os.listdir(path)
    n=0
    now1= date.today() 
    for i in range(30):   
        date_name_4G = 'eNodebList_4G({}).xlsx'.format(now1)
        date_name_5G = 'eNodebList_5G({}).xlsx'.format(now1)
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

def find_last_record():
    list1=['']
    path= './results(analysis)'
    date_list = os.listdir(path)
    now= date.today() 
    now1 = now - datetime.timedelta(days=1)
    for i in range(30):   
        antenna_df = 'bts_antenna_switch({}).xlsx'.format(now1)
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
        pcb_df = '天線庫存({}).xlsx'.format(now)
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
assets_df = pd.read_excel(file1_path, sheet_name = '基地台天線')

# -------讀取 nmoss 基地台資料庫 (nmoss_df) -------------

# nmoss_files =['eNodebList_4G(2021-05-11).xlsx','eNodebList_5G(2021-05-11).xlsx']
nmoss_files = find_earlier_eNodebList() # 使用 def 

# ---------寫入資料庫  ---------------------------------
today = date.today() 
writer =pd.ExcelWriter('./results(analysis)/bts_antenna_switch({}).xlsx'.format(today))   

#---------調整 assets_df 格式 -----------
assets_df = assets_df[assets_df['財產名稱'].isin(['２Ｇ及３Ｇ行動電話室外涵蓋型共用天線系統','4G/3G/2G行動通信室內涵蓋天線','4G/3G/2G行動通信室外涵蓋天線','5G/4G/3G室外涵蓋天線'])]
assets_df['廠牌'] = assets_df['廠牌'].str.capitalize()
assets_inuse_df = assets_df[(assets_df['設備狀態']=='使用中') & (assets_df['數量'] == 1)]  
assets_inuse_df = assets_inuse_df[assets_inuse_df['編號'].str.contains('N|L')]  #

assets_spare_df = assets_df[(assets_df['設備狀態']=='備援/備用') & (assets_df['數量']== 1)]
assets_stop_df = assets_df[(assets_df['設備狀態']=='停用') & (assets_df['數量']== 1)]
assets_loc_df = assets_df[(assets_df['設備狀態']=='佔位置') & (assets_df['數量']== 1)]
assets_loss_df = assets_df[(assets_df['設備狀態']=='已遺失') & (assets_df['數量']== 1)]

assets_inuse_df.to_excel(writer,sheet_name ='assets(使用中)',index=False)
assets_spare_df.to_excel(writer,sheet_name ='assets(備用)',index=False)
assets_stop_df.to_excel(writer,sheet_name ='assets(停用)',index=False)
assets_loc_df.to_excel(writer,sheet_name ='assets(佔位)',index=False)
assets_loss_df.to_excel(writer,sheet_name ='assets(遺失)',index=False)

#---------調整 assets_df 格式 -----------

assets_data_df = assets_inuse_df[['編號','基地台名稱','廠牌','型式/號','財產編號','異動者']]
temp_df = assets_inuse_df['廠牌'] + '_'+ assets_inuse_df['型式/號'].astype(str)
assets_data_df = assets_data_df.drop(['廠牌','型式/號'],axis =1)
assets_data_df.insert(2,'天線型號(assets)',temp_df)
#assets_data_df['天線型號(assets)'] = assets_data_df['天線型號(assets)'].str.replace('kathrein','Kathrein').str.replace('Gamma NU_GN-MB4P11','GAMMANU_GN-MB4P11').str.replace('Gamma NU_GN-MB4P8','GAMMANU_GN-MB4P8') 
assets_data_df.to_excel(writer,sheet_name ='assets',index=False)  ###

#---------調整 nmoss_df 格式 ------------
small_bts_exception =['601013L','601016L','601018L','601020L','601023L','601028L','601029L','601032L','601033L','601035L','601037L',
                     '601038L','601039L','601040L','601045L','601046L','601047L','601048L','601049L','601050L','601051L','601055L',
                     '601056L','601057L','601058L','601059L','601069L','601070L','601071L','601072L','601073L','601074L','601077L',
                     '601079L','601080L','601087L','601097L','601099L','601103L','601119L','605173L','605259L','601000L','601014L',
                     '601017L','601021L','601022L','601025L','601026L','601027L','601030L','601031L','601041L','601042L','601043L',
                     '601044L','601053L','601062L','601066L','601075L','601081L','601088L','601090L','601091L','601092L','601093L',
                     '601094L','601095L','601096L','601107L','601114L','602290L','605113L','605138L','605328L','606894L','607039L',
                     '607041L']
n = 0
for filename in nmoss_files:
    begin_df = pd.read_excel('./data_eNodebList/'+ filename,sheet_name = 0)
    list_colname = list(begin_df.head())
    if n==0:
        begin_df.rename(columns = {list_colname[2]:'編號'},inplace = True)
        temp4G_df = begin_df
        n = n + 1
    else:
        begin_df.rename(columns = {list_colname[1]:'編號'},inplace = True)
        temp5G_df =begin_df
        
temp4G_df = temp4G_df[['編號','扇區編號(sectorno)','基地台名稱(BName)','天線廠牌1(AntennaBrand1)','天線型號1(AntennaType1)','天線廠牌2(AntennaBrand2)','天線型號2(AntennaType2)','天線廠牌3(AntennaBrand3)','天線型號3(AntennaType3)']] 

smallbts_with_twoant = temp4G_df[temp4G_df['編號'].isin(small_bts_exception)]
temp4G_df = temp4G_df.drop(temp4G_df.loc[temp4G_df['編號'].isin(small_bts_exception)].index)
temp4G_df = temp4G_df.drop_duplicates()
temp4G_df = pd.concat([smallbts_with_twoant, temp4G_df], ignore_index=True)
temp4G_df = temp4G_df.reset_index(drop = True)
temp4G_df = temp4G_df.drop(['扇區編號(sectorno)'], axis=1)

#--調整 1支天線 tri-sector (splitter) 2 port [Commscope V360QS-C3-3XR]--
index1 = temp4G_df.loc[temp4G_df['天線型號1(AntennaType1)']=='V360QS-C3-3XR'].index
temp4G_df.loc[index1,'天線廠牌2(AntennaBrand2)':'天線型號3(AntennaType3)'] = np.NaN

#---------特例調整 (612299L: 用 Andrew_3X-V65A-3XR 天線，但卻於外部作 splitter)--
index5 = temp4G_df.loc[(temp4G_df['編號']=='612299L') & (temp4G_df['天線型號1(AntennaType1)']=='3X-V65A-3XR')].index
temp4G_df.loc[index5,'天線廠牌2(AntennaBrand2)':'天線型號3(AntennaType3)'] = np.NaN

#----調整 tri-sector天線 (3方向但只有一支天線 6 port)[BROADRADIO_LLLOX306R-D],[Andrew_3X-V65A-3XR],[COMMSCOPE_NNNOX310R]--
except_trisector_df = temp4G_df[temp4G_df['天線型號1(AntennaType1)'].isin(['LLLOX306R-D','3X-V65A-3XR','NNNOX310R'])]                       
except_trisector_df = except_trisector_df.drop_duplicates() # except_trisector_df 後面必須加回去
index2 = temp4G_df.loc[temp4G_df['天線型號1(AntennaType1)'].isin(['LLLOX306R-D','3X-V65A-3XR','NNNOX310R'])].index
temp4G_df.drop(index2,axis = 0,inplace = True)
temp4G_df = pd.concat([temp4G_df,except_trisector_df])
temp4G_df = temp4G_df.reset_index(drop = True)
#-------------------------------------------------------------------------



temp5G_df = temp5G_df[['編號','細胞編號(CellID)','基地台名稱(BName)','天線廠牌1(AntennaBrand1)','天線型號1(AntennaType1)','天線廠牌2(AntennaBrand2)','天線型號2(AntennaType2)','天線廠牌3(AntennaBrand3)','天線型號3(AntennaType3)']]   
temp_df['check'] = temp5G_df['細胞編號(CellID)'].map(choice_char)
temp5G_df = temp5G_df.drop(['細胞編號(CellID)'], axis =1)
temp5G_df['check'] = temp_df['check']
temp5G_df = temp5G_df.drop_duplicates()
temp5G_df = temp5G_df.drop(['check'], axis =1)
# temp5G_df.to_excel(writer,sheet_name ='temp5G',index=False)  #

nmoss_orig_df = pd.concat([temp4G_df, temp5G_df])
nmoss_orig_df = nmoss_orig_df.reset_index(drop = True)

nmoss_ante1_df = nmoss_orig_df[['編號','基地台名稱(BName)','天線廠牌1(AntennaBrand1)','天線型號1(AntennaType1)']] 
nmoss_ante1_df = nmoss_ante1_df.dropna(axis=0)
nmoss_ante1_df['天線型號(nmoss)'] = nmoss_ante1_df['天線廠牌1(AntennaBrand1)'].str.capitalize() + '_' + nmoss_ante1_df['天線型號1(AntennaType1)'].astype(str)
nmoss_ante1_df = nmoss_ante1_df.drop(['天線廠牌1(AntennaBrand1)','天線型號1(AntennaType1)'],axis =1)

nmoss_ante2_df = nmoss_orig_df[['編號','基地台名稱(BName)','天線廠牌2(AntennaBrand2)','天線型號2(AntennaType2)']] 
nmoss_ante2_df = nmoss_ante2_df.dropna(axis=0)
nmoss_ante2_df['天線型號(nmoss)'] = nmoss_ante2_df['天線廠牌2(AntennaBrand2)'] .str.capitalize()+ '_' + nmoss_ante2_df['天線型號2(AntennaType2)'].astype(str)
nmoss_ante2_df = nmoss_ante2_df.drop(['天線廠牌2(AntennaBrand2)','天線型號2(AntennaType2)'],axis =1)

nmoss_ante3_df = nmoss_orig_df[['編號','基地台名稱(BName)','天線廠牌3(AntennaBrand3)','天線型號3(AntennaType3)']] 
nmoss_ante3_df = nmoss_ante3_df.dropna(axis=0)
nmoss_ante3_df['天線型號(nmoss)'] = nmoss_ante3_df['天線廠牌3(AntennaBrand3)'].str.capitalize() + '_' + nmoss_ante3_df['天線型號3(AntennaType3)'].astype(str)
nmoss_ante3_df = nmoss_ante3_df.drop(['天線廠牌3(AntennaBrand3)','天線型號3(AntennaType3)'],axis =1)


nmoss_data_df = pd.concat([nmoss_ante1_df, nmoss_ante2_df, nmoss_ante3_df])
nmoss_data_df = nmoss_data_df.reset_index(drop = True)
#
nmoss_data_df.sort_values(by = ['編號','天線型號(nmoss)'],inplace = True)
nmoss_data_df = nmoss_data_df.reset_index(drop = True)
nmoss_data_df = nmoss_data_df.drop(nmoss_data_df.loc[nmoss_data_df['天線型號(nmoss)'].str.contains(pat ='Nokia_')].index)
nmoss_data_df.to_excel(writer,sheet_name ='nmoss',index=False)  ##
nmoss_data_df['天線型號(nmoss)'] = nmoss_data_df['天線型號(nmoss)'].str.replace(')','').str.replace('(','').str.replace('UT45-N2','UT45').str.replace('不可申請證照','').str.replace('1710-2690','')
#nmoss_data_df['天線型號(nmoss)'] = nmoss_data_df['天線型號(nmoss)'].str.replace('Gammanu','Gamma nu')

#---------合併 (assets) (nmoss) 兩資料庫內容 ------------
full_df = pd.concat([assets_data_df, nmoss_data_df])
full_df = full_df.reset_index(drop = True)
full_df = full_df[['編號','基地台名稱(BName)','天線型號(nmoss)','天線型號(assets)','財產編號','異動者']]
full_df.sort_values(by = ['編號','天線型號(nmoss)','天線型號(assets)'],inplace = True)
full_df = full_df.reset_index(drop = True)
# full_df.to_excel(writer,sheet_name ='combined',index=False)

# --------比較 (assets_df) (netact_df)兩資料庫內容 ----------- 
bts_id = list(full_df['編號'].unique())
for i in bts_id:
    nmoss_tmp_df = full_df[(full_df['編號']== i)&(full_df['天線型號(assets)'].isnull())]
    check1 = list(nmoss_tmp_df['天線型號(nmoss)'].str.replace('Gammanu','Gamma nu').str.replace('Andrew_DBXLH-6565A-VTM','Commscope_DBXLH-6565A-VTM').str.replace('Andrew_3X-V65A-3XR','Commscope_3X-V65A-3XR').
                 str.replace('Argus_NOX310R','Commscope_NOX310R').str.replace('LLPX202F0','LPX202F').str.replace('Andrew_HBX-6516DS-VTM','Commscope_HBX-6516DS-VTM').str.replace('Argus_NNNOX310R','Commscope_NNNOX310R').str.replace('Andrew_DBXDH-6565B-VTM','Commscope_DBXDH-6565B-VTM'))
    check_1 = [s.upper() for s in check1]
    check_1.sort() 
    
    assets_tmp_df = full_df[(full_df['編號']== i)&(full_df['天線型號(nmoss)'].isnull())]
    check2 = list(assets_tmp_df['天線型號(assets)'].str.replace('Andrew_DBXLH-6565A-VTM','Commscope_DBXLH-6565A-VTM').str.replace('Andrew_3X-V65A-3XR','CommScope_3X-V65A-3XR').
                 str.replace('Argus_NOX310R','Commscope_NOX310R').str.replace('LLPX202F0','LPX202F').str.replace('Andrew_HBX-6516DS-VTM','Commscope_HBX-6516DS-VTM').str.replace('Argus_NNNOX310R','Commscope_NNNOX310R').str.replace('Andrew_DBXDH-6565B-VTM','Commscope_DBXDH-6565B-VTM'))
     
    check_2 =[s.upper() for s in check2] 
    check_2.sort()  
    
    if check_1 != check_2 :
        index1 = full_df.loc[full_df['編號']==i].index 
        full_df.loc[index1 ,'check'] = 'X'
        
        del_a = []
        a = check_2
        b = check_1
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
        
full_df.to_excel(writer,sheet_name ='天線調整',index = False)
worksheet = writer.sheets['天線調整']
worksheet.set_column("B:E",26)


# 成績值比較 先選以前值
last_record =find_last_record() # 使用 def Macro
last_record_df = pd.read_excel('./results(analysis)/'+ last_record[0],sheet_name = '完成數',dtype= {'日期': str})
#last_record_df['日期'] = last_record_df['日期'].dt.strftime('%Y-%m-%d')

grade_df = full_df.loc[full_df.check=='O'].copy()
grade_df['異動者'].fillna(method='bfill',inplace=True)
grade_df.drop(['基地台名稱(BName)','天線型號(nmoss)','天線型號(assets)','財產編號','check'],axis =1,inplace=True)
grade_df = grade_df.drop_duplicates()
grade1_df = pd.DataFrame(grade_df['異動者'].value_counts()).T
grade1_df['日期']=str(today)
grade1_df.insert(0, '日期', grade1_df.pop('日期'))

result = pd.concat([last_record_df, grade1_df],ignore_index= True)
result = result.reset_index(drop = True)

result.tail(15).to_excel(writer,sheet_name ='完成數',index = False)
worksheet = writer.sheets['完成數']
worksheet.set_column("A:A",12)

#============建立統計表(begin)===================#

assets_inuse_df['廠牌'] = assets_inuse_df['廠牌'].str.capitalize()
assets_inuse_df['廠牌'] = assets_inuse_df['廠牌'].str.replace('Argus','Commscope')
assets_inuse_df['型式/號'] = assets_inuse_df['型式/號'].astype(str).str.upper()
assets_inuse_df['assets_使用中'] = assets_inuse_df['廠牌']+'_'+assets_inuse_df['型式/號'].astype(str)
assets_inuse_df['assets_使用中'] = assets_inuse_df['assets_使用中'].str.replace('Andrew_DBXDH-6565B-VTM','Commscope_DBXDH-6565B-VTM').str.replace('Andrew_3X-V65A-3XR','Commscope_3X-V65A-3XR').str.replace('Andrew_DBXLH-6565A-VTM','Commscope_DBXLH-6565A-VTM')
assets_inuse_count = assets_inuse_df['assets_使用中'].value_counts()
assets_inuse_stic = pd.DataFrame(assets_inuse_count)
assets_inuse_stic.index.name="天線型式"

assets_spare1_df = assets_spare_df.copy()
assets_spare1_df['廠牌'] = assets_spare1_df['廠牌'].str.capitalize()
assets_spare1_df['廠牌'] = assets_spare1_df['廠牌'].str.replace('Argus','Commscope')
assets_spare1_df['型式/號'] = assets_spare1_df['型式/號'].astype(str).str.upper()
assets_spare1_df['assets_備用'] = assets_spare1_df['廠牌']+'_'+assets_spare1_df['型式/號'].astype(str)
assets_spare1_df['assets_備用'] = assets_spare1_df['assets_備用'].str.replace('Andrew_DBXDH-6565B-VTM','Commscope_DBXDH-6565B-VTM').str.replace('Andrew_3X-V65A-3XR','Commscope_3X-V65A-3XR').str.replace('Andrew_DBXLH-6565A-VTM','Commscope_DBXLH-6565A-VTM')
assets_spare_count = assets_spare1_df['assets_備用'].value_counts()
assets_spare_stic = pd.DataFrame(assets_spare_count)
assets_spare_stic.index.name="天線型式"

assets_spare_north =  assets_spare1_df[assets_spare1_df['使用單位']=='北嘉義基維股']
assets_spare_north_cnt = assets_spare_north['assets_備用'].value_counts()
assets_spare_norstic = pd.DataFrame(assets_spare_north_cnt)
assets_spare_norstic.rename(columns={'assets_備用':'北基備用'},inplace=True)
assets_spare_norstic.index.name="天線型式"

assets_spare_south =  assets_spare1_df[assets_spare1_df['使用單位']=='南嘉義基維股']
assets_spare_south_cnt = assets_spare_south['assets_備用'].value_counts()
assets_spare_soustic = pd.DataFrame(assets_spare_south_cnt)
assets_spare_soustic.rename(columns={'assets_備用':'南基備用'},inplace=True)
assets_spare_soustic.index.name="天線型式"

assets_spare_qual =  assets_spare1_df[assets_spare1_df['使用單位']=='嘉義品改股']
assets_spare_qual_cnt = assets_spare_qual['assets_備用'].value_counts()
assets_spare_qualstic = pd.DataFrame(assets_spare_qual_cnt)
assets_spare_qualstic.rename(columns={'assets_備用':'品改備用'},inplace=True)
assets_spare_qualstic.index.name="天線型式"

assets_stop1_df = assets_stop_df.copy()
assets_stop1_df['廠牌'] = assets_stop1_df['廠牌'].str.capitalize()
assets_stop1_df['廠牌'] = assets_stop1_df['廠牌'].str.replace('Argus','Commscope')
assets_stop1_df['型式/號'] = assets_stop1_df['型式/號'].astype(str).str.upper()
assets_stop1_df['assets_停用'] = assets_stop1_df['廠牌']+'_'+assets_stop1_df['型式/號'].astype(str)
assets_stop1_df['assets_停用'] = assets_stop1_df['assets_停用'].str.replace('Andrew_DBXDH-6565B-VTM','Commscope_DBXDH-6565B-VTM').str.replace('Andrew_3X-V65A-3XR','Commscope_3X-V65A-3XR').str.replace('Andrew_DBXLH-6565A-VTM','Commscope_DBXLH-6565A-VTM')
assets_stop_count  = assets_stop1_df['assets_停用'].value_counts()
assets_stop_stic = pd.DataFrame(assets_stop_count)
assets_stop_stic.index.name="天線型式"

assets_stop_north =  assets_stop1_df[assets_stop1_df['使用單位']=='北嘉義基維股']
assets_stop_north_cnt = assets_stop_north['assets_停用'].value_counts()
assets_stop_norstic = pd.DataFrame(assets_stop_north_cnt)
assets_stop_norstic.rename(columns={'assets_停用':'北基停用'},inplace=True)
assets_stop_norstic.index.name="天線型式"

assets_stop_south =  assets_stop1_df[assets_stop1_df['使用單位']=='南嘉義基維股']
assets_stop_south_cnt = assets_stop_south['assets_停用'].value_counts()
assets_stop_soustic = pd.DataFrame(assets_stop_south_cnt)
assets_stop_soustic.rename(columns={'assets_停用':'南基停用'},inplace=True)
assets_stop_soustic.index.name="天線型式"

nmoss_data_df['天線型號(nmoss)'] = nmoss_data_df['天線型號(nmoss)'].str.replace('Andrew_DBXLH-6565A-VTM','Commscope_DBXLH-6565A-VTM').str.replace('Gammanu','Gamma nu')
nmoss_data_df['天線型號(nmoss)'] = nmoss_data_df['天線型號(nmoss)'].str.replace('LLPX202F0','LPX202F').str.replace('Andrew_DBXDH-6565B-VTM','Commscope_DBXDH-6565B-VTM').str.replace('Andrew_3X-V65A-3XR','Commscope_3X-V65A-3XR').str.replace('AARC','Aarc').str.replace('CommScope','Commscope').str.replace('COMMSCOPE','Commscope').str.replace('BROADRADIO','Broadradio')
nmoss_data_df['天線型號(nmoss)'] = nmoss_data_df['天線型號(nmoss)'].str.replace('Commscope_HBX-6516DS-VTM','Andrew_HBX-6516DS-VTM')

nmoss_data_df_count = nmoss_data_df['天線型號(nmoss)'].value_counts()
nmoss_data_df_stic = pd.DataFrame(nmoss_data_df_count)
nmoss_data_df_stic.rename(columns={'天線型號(nmoss)':'nmoss_使用中'},inplace=True)
nmoss_data_df_stic.index.name="天線型式"


statistics_df = assets_inuse_stic.join(assets_spare_stic,how='outer')
statistics_df = statistics_df.join(assets_spare_norstic,how='outer')
statistics_df = statistics_df.join(assets_spare_soustic,how='outer')
statistics_df = statistics_df.join(assets_spare_qualstic,how='outer')
statistics_df = statistics_df.join(assets_stop_stic,how='outer')
statistics_df = statistics_df.join(assets_stop_norstic,how='outer')
statistics_df = statistics_df.join(assets_stop_soustic,how='outer')
statistics_df = statistics_df.join(nmoss_data_df_stic,how='outer')
statistics_df = statistics_df.fillna(0)
statistics_df[['assets_使用中','assets_備用','北基備用','南基備用','品改備用','assets_停用','北基停用','南基停用','nmoss_使用中']] = statistics_df[['assets_使用中','assets_備用','北基備用','南基備用','品改備用','assets_停用','北基停用','南基停用','nmoss_使用中']].astype(int)


#============讀取電路板庫存量(in stock)================#
file1_name = find_earlier_in_stock() # 使用 def 
file1_path = "./data_in_stock/{}".format(file1_name[0])
in_stock_df = pd.read_excel(file1_path, sheet_name = '統計')
in_stock_df.fillna(value=0, inplace=True)
in_stock_df['庫存_嘉義']= in_stock_df['庫存_北基']+in_stock_df['庫存_南基']+in_stock_df['庫存_品改']
in_stock_df.set_index("天線型式" , inplace=True)

statistics_df = statistics_df.join(in_stock_df,how='outer')
statistics_df = statistics_df[['assets_使用中','assets_備用','北基備用','南基備用','品改備用','assets_停用','北基停用','南基停用','nmoss_使用中','庫存_嘉義','庫存_北基','庫存_南基','庫存_品改']]
statistics_df.fillna(value=0, inplace=True)
statistics_df['財編缺額'] = statistics_df['nmoss_使用中'] + statistics_df['庫存_嘉義'] - statistics_df['assets_使用中'] - statistics_df['assets_備用'] - statistics_df['assets_停用'] 

statistics_df = statistics_df.reset_index()
stat_style = statistics_df.style.applymap(lambda x: 'background-color:#ADD8E6', subset=["北基備用"])     .applymap(lambda x: 'background-color:#ADD8E6', subset=["庫存_北基"])     .applymap(lambda x: 'background-color:#FFFF74', subset=["assets_停用"])     .applymap(lambda x: 'background-color:#FFFF74', subset=["北基停用"])     .applymap(lambda x: 'background-color:#FFFF74', subset=["南基停用"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["南基備用"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_南基"]) 

stat_style.to_excel(writer,sheet_name ='統計表',index = False)
worksheet = writer.sheets['統計表']
worksheet.set_column("A:A",34)
worksheet.set_column("B:C",15)
worksheet.set_column("D:F",11)
worksheet.set_column("G:G",13)
worksheet.set_column("H:I",11)
worksheet.set_column("J:J",15)
worksheet.set_column("K:O",11)

#----------------統計表 大於 > 0  ---------------------#
# 1.assets_使用中 < nmoss_使用中 ，assets_備用&assets_停用 !=0  提出
# 2.nmoss_使用中 == nmoss_使用中 
# 備用 < 庫存， 停用 !=0 提出
# 備用 > 庫存  提出
# 3.assets_使用中 > nmoss_使用中 提出

A = statistics_df['財編缺額'] > 0
stat_great_0_df = statistics_df[A]

B = stat_great_0_df['assets_使用中'] < stat_great_0_df['nmoss_使用中']
C = stat_great_0_df['assets_備用'] != 0
D = stat_great_0_df['assets_停用'] != 0
E = B&(C|D)

F = stat_great_0_df['assets_使用中'] == stat_great_0_df['nmoss_使用中']
G = stat_great_0_df['南基備用'] < stat_great_0_df['庫存_南基']
H = stat_great_0_df['南基停用'] != 0 
I = stat_great_0_df['北基備用'] < stat_great_0_df['庫存_北基']
J = stat_great_0_df['北基停用'] != 0
K = stat_great_0_df['南基備用']> stat_great_0_df['庫存_南基']
L = stat_great_0_df['北基備用']> stat_great_0_df['庫存_北基']
M = F&((G&H)|(I&J)|K|L)

N = stat_great_0_df['assets_使用中'] > stat_great_0_df['nmoss_使用中']

stat_great_1_df = stat_great_0_df[E|M|N]

great_style = stat_great_1_df.style.applymap(lambda x: 'background-color:#ADD8E6', subset=["assets_備用"])       .applymap(lambda x: 'background-color:#ADD8E6', subset=["assets_停用"])       .applymap(lambda x: 'background-color:#FFFF74', subset=["assets_使用中"])       .applymap(lambda x: 'background-color:#FFFF74', subset=["nmoss_使用中"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["北基備用"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["南基備用"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_北基"])       .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_南基"])

great_style.to_excel(writer,sheet_name ='缺額>0修正',index = False)
worksheet = writer.sheets['缺額>0修正']
worksheet.set_column("A:A",34)
worksheet.set_column("B:C",15)
worksheet.set_column("D:F",11)
worksheet.set_column("G:G",13)
worksheet.set_column("H:I",11)
worksheet.set_column("J:J",15)
worksheet.set_column("K:O",11)

#----------------統計表 小於 <=0  ----------------------#
# 1.assets_使用中 < nmoss_使用中  提出
# 2.assets_使用中 == nmoss_使用中 
# 備用 < 庫存， 停用 !=0 提出
# 備用 > 庫存  提出
# 3.assets_使用中 > nmoss_使用中 提出


A = statistics_df['財編缺額'] <= 0

stat_less_0_df = statistics_df[A]

B = stat_less_0_df['assets_使用中'] < stat_less_0_df['nmoss_使用中']

C = stat_less_0_df['assets_使用中'] == stat_less_0_df['nmoss_使用中']
D = stat_less_0_df['南基備用'] < stat_less_0_df['庫存_南基']
E = stat_less_0_df['南基停用'] != 0 
F = stat_less_0_df['南基備用'] >  stat_less_0_df['庫存_南基']
G = stat_less_0_df['北基備用'] < stat_less_0_df['庫存_北基']
H = stat_less_0_df['北基停用'] != 0 
I = stat_less_0_df['北基備用'] >  stat_less_0_df['庫存_北基']
J = C&((D&E)|F|(G&H)|I)

K = stat_less_0_df['assets_使用中'] > stat_less_0_df['nmoss_使用中']

stat_less_1_df = stat_less_0_df[B|J|K]

less_style = stat_less_1_df.style.applymap(lambda x: 'background-color:#ADD8E6', subset=["assets_使用中"])     .applymap(lambda x: 'background-color:#ADD8E6', subset=["nmoss_使用中"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["北基備用"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["南基備用"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_北基"])     .applymap(lambda x: 'background-color:#E6C3C3', subset=["庫存_南基"]) 

less_style.to_excel(writer,sheet_name ='缺額<=0修正',index = False)
worksheet = writer.sheets['缺額<=0修正']
worksheet.set_column("A:A",34)
worksheet.set_column("B:C",15)
worksheet.set_column("D:F",11)
worksheet.set_column("G:G",13)
worksheet.set_column("H:I",11)
worksheet.set_column("J:J",15)
worksheet.set_column("K:O",11)

#===============建立統計表(end)========================#
writer.save()




# In[ ]:





# In[ ]:




