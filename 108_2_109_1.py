import pandas as pd
import pyodbc
# Some other example server values are
# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port

server = ''
database = ''
username = ''
password = ''
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

cursor.execute("SELECT @@version;")
row = cursor.fetchone()
while row:
    print(row[0])
    row = cursor.fetchone()
cursor = cnxn.cursor()
#course_info
query1 = "SELECT * FROM thuir_studscore_attr WHERE SETYEAR = '109' AND SETTERM = '2' AND type_name= '日間學士班';"
df_109_2 = pd.read_sql(query1, cnxn)
df_109_2.to_excel (r'D:\AI_學生效益分析108_2開始\df_109_2.xlsx', index = False, header=True)

# # query2 = "SELECT * FROM THUIR_STUDSCORE_ATTR WHERE SETYEAR = '109' AND SETTERM = '1' AND type_name= '日間學士班';"
# # df_109_1 = pd.read_sql(query2, cnxn)
query3 = "SELECT * FROM thuir_studscore_attr WHERE SETYEAR = '109' AND SETTERM = '1' AND type_name= '日間學士班';"
df_109_1_new = pd.read_sql(query3, cnxn)

count = df_109_1_new['CURR_ATTR'].value_counts()
print(count)


query2 = "SELECT * FROM thuir_studscore_attr WHERE SETYEAR = '108' AND SETTERM = '2' AND type_name= '日間學士班';"
df_108_2_new = pd.read_sql(query2, cnxn)

count = df_108_2_new['CURR_ATTR'].value_counts()
print(count)

count = df_109_2['CURR_ATTR'].value_counts()
print(count)
