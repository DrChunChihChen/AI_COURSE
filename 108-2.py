import pandas as pd
import pyodbc
# Some other example server values are
# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port

server = '140.128.97.85'
database = 'thuir'
username = 'thuir'
password = 'thuir'
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

cursor.execute("SELECT @@version;")
row = cursor.fetchone()
while row:
    print(row[0])
    row = cursor.fetchone()
cursor = cnxn.cursor()
#course_info


query2 = "SELECT * FROM thuir_studscore_attr WHERE SETYEAR = '108' AND SETTERM = '1' AND type_name= '日間學士班';"
df_108_2_new = pd.read_sql(query2, cnxn)

# count = df_108_2_new['CURR_ATTR'].value_counts()
# print(count)


#------------------------留下有AI相關課程的學號---------------------------#
# # df109_1= df_109_1_new[df_109_1_new['CURR_ATTR'].isin(['全校性AI課程'])]
# df108_2= df_108_2_new[df_108_2_new['CURR_ATTR'].isin(['全校性AI課程'])]
# # df109_1= df_109_1[df_109_1['CURR_ATTR'].isin(['全校性AI課程', 'AI計畫L1','AI計畫L2','程式設計-Level2','程式設計-Level1','程式設計-Level3','創新學院-雲創'])]
# # df108_2= df_108_2[df_108_2['CURR_ATTR'].isin(['全校性AI課程', 'AI計畫L1','AI計畫L2','程式設計-Level2','程式設計-Level1','程式設計-Level3','創新學院-雲創'])]
#
# # new_1091 = df_109_1_new[['STUD_NO']]
# new_1082 = df_108_2_new[['STUD_NO']]
# # new_1091 = df109_1[['STUD_NO']]
# # new_1082 = df108_2[['STUD_NO']]
#
# # new_1091 = new_1091.drop_duplicates('STUD_NO', keep='last')
# new_1082 = new_1082.drop_duplicates('STUD_NO', keep='last')
# # new_1091 = new_1091.drop_duplicates('STUD_NO', keep='last')
# # new_1082 = new_1082.drop_duplicates('STUD_NO', keep='last')

# #-----------------------------------------108-2---------------------------------------------------------------------------------------
# df108_2＿修課 = df_108_2_new[df_108_2_new['STUD_NO'].isin(new_1082['STUD_NO'])]
# df108_2＿修課.to_excel (r'D:\AI_學生效益分析108_2開始\df108_2＿修課.xlsx', index = False, header=True)


df108_2_AI生修課狀況 = df_108_2_new[~df_108_2_new['MAJR2_NAME'].isin(['大一大二體育', '軍訓一','大一英文','通識課程:自然領域',
                                                          '通識課程:人文領域','通識課程:社會領域','大二英文','通識課程:文明與經典'
                                                             ,'通識課程:多元與與議題導向','選修英語','通識課程:領導與倫理'])]
df108_2_AI生修課狀況['MAJR_FULL_NAME_2'] = df108_2_AI生修課狀況['MAJR_FULL_NAME'].replace({'化學系化學生物組':'化學系',
                                                                                 '化學系化學組':'化學系',
                                                                                 '生命科學系生物醫學組':'生命科學系',
                                                                                 '生命科學系生態暨生物多樣性組':'生命科學系',
                                                                                 '政治學系政治理論組':'政治學系',
                                                                                 '政治學系國際關係組':'政治學系',
                                                                                 '經濟學系一般經濟組':'經濟學系',
                                                                                 '經濟學系產業經濟組':'經濟學系',
                                                                                 '資訊工程學系軟體工程組':'資訊工程學系',
                                                                                 '資訊工程學系資電工程組':'資訊工程學系',
                                                                                 '資訊工程學系數位創意組':'資訊工程學系',
                                                                                 '電機工程學系IC設計與無線通訊組':'電機工程學系',
                                                                                 '電機工程學系奈米電子與能源技術組':'電機工程學系',
                                                                                 '應用物理學系光電組':'應用物理學',
                                                                                 '應用物理學系材料及奈米科技組':'應用物理學'})

df108_2_AI生修課狀況 ['跨域'] = (df108_2_AI生修課狀況 ['MAJR_FULL_NAME_2'] == df108_2_AI生修課狀況 ['MAJR1_NAME'])
df108_2_AI生修課狀況.to_excel (r'D:\AI_學生效益分析108_2開始\df108_2_AI生修課狀況.xlsx', index = False, header=True)

#-----------------------------------全校性AI課程+乾淨的修課紀錄------------------------------------

# df108_2_AI生修課狀況s = df108_2_AI生修課狀況.loc[df108_2_AI生修課狀況['CURR_ATTR']=='全校性AI課程',:]
# df108_2_AI生修課狀況t = df108_2_AI生修課狀況[~df108_2_AI生修課狀況['CURR_ATTR'].isin(['全校性AI課程'])]
#
# df108_2_AI生修課狀況t = df108_2_AI生修課狀況t.drop_duplicates(subset=['STUD_NO', 'CURR_CODE'], keep='first')
# listdf = [df108_2_AI生修課狀況t, df108_2_AI生修課狀況s]
# df108_2_AI生修課狀況_new = pd.concat(listdf)
# df108_2_AI生修課狀況_new.sort_values(by='STUD_NO', inplace=True)
# df108_2_AI生修課狀況_new = df108_2_AI生修課狀況_new.drop_duplicates(subset=['STUD_NO', 'CURR_CODE'], keep='first')
# df108_2_AI生修課狀況_new.to_excel (r'D:\AI_學生效益分析108_2開始\df108_2_AI生修課狀況_new .xlsx', index = False, header=True)

#---------------------------------------------------計算跨域 各類修課成績---------------------------------------------------------------------------

crosstab_df1082 = pd.crosstab(df108_2_AI生修課狀況.STUD_NO,df108_2_AI生修課狀況.跨域).apply(lambda r: r/r.sum(), axis=1)
crosstab_df1082.to_excel (r'D:\AI_學生效益分析108_2開始\df108_2_AI生修課狀況crosstab_df.xlsx', index = False, header=True)
crosstab_df_跨域1082_new = crosstab_df1082.reset_index(level=0, inplace=False)

import numpy as np
df108_2_AI生修課狀況['SCORE_2'] = df108_2_AI生修課狀況['SCORE'].replace({
                                                                                 '通過':np.nan,
                                                                                 '未過':np.nan,
                                                                                 '抵免':np.nan,})
df108_2_AI生修課狀況['SCORE_2'] = df108_2_AI生修課狀況['SCORE_2'].astype('float')
crosstab_df_成績1082 = pd.crosstab(index = df108_2_AI生修課狀況.STUD_NO, columns =df108_2_AI生修課狀況.CURR_SEL, values=df108_2_AI生修課狀況.SCORE_2, aggfunc=np.nanmean)

crosstab_df_成績1082_new = crosstab_df_成績1082.reset_index(level=0, inplace=False)

df108_2_AI生修課狀況_new = pd.merge(df108_2_AI生修課狀況, crosstab_df_跨域1082_new, left_on="STUD_NO", right_on="STUD_NO")
df108_2_AI生修課狀況_new = pd.merge(df108_2_AI生修課狀況, crosstab_df_跨域1082_new , left_on="STUD_NO", right_on="STUD_NO")

# df108_2_AI生修課狀況_new.to_excel (r'D:\AI_學生效益分析108_2開始\df108_1_全校生修課狀況_new_跨域_成績.xlsx', index = False, header=True)


df108_2_AI生修課狀況_new1 = df108_2_AI生修課狀況_new.drop_duplicates(subset=['STUD_NO', 'CURR_CODE'], keep='first')
df108_2_AI生修課狀況_new1.to_excel (r'D:\AI_學生效益分析108_2開始\df108_11_全校生修課狀況_new_跨域_成績.xlsx', index = False, header=True)