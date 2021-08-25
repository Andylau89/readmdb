# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import pypyodbc

p_path = r'C:\Users\heliu001\Desktop\python\mdb\test.mdb'
p_path=input("请输入路径:")
connStr = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+ p_path
# print(connStr)


conn = pypyodbc.win_connect_mdb(connStr)  # 链接数据库
cur = conn.cursor()  # 创建游标
sql = "SELECT * FROM " + 'srv'  # 取表 ActualValues_T
cur.execute(sql)
alldata = cur.fetchall()  # 取 ActualValues_T 所有数据


total_rows = len(alldata)
total_cols = len(alldata[0])
print("****************Begin to process\"表:srv\"****************")
print("\"表:%s\"总行数 = %d" % ('srv', total_rows))
print("\"表:%s\"总列数 = %d" % ('srv', total_cols))
print(type(alldata))
# print(alldata)
conn.close()  # 关闭数据库
for i in range(len(alldata)):
    print(alldata[i])
