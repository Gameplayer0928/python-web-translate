# coding=utf-8
'''
Created on 2018年3月6日
 @author: Gameplayer0928 Qi Gao

#
#    This file is part of exceltrans v0.1.
#
#    exceltrans v0.1  is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    exceltrans v0.1 is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with exceltrans v0.1.  If not, see <http://www.gnu.org/licenses/>.
#
#    Copyright 2018, 2019, 2020 Qi Gao
# 

'''

''' this program is translate chinese to english with YouDao web. 
    create Table in MySQL database,
    read Excel file and put data to MySQL database,
    read data from database an put to YouDao web for translate,
    add 'english' data type in database and put translated data into database in 'english'. 
 '''

import xlrd
import xlwt
import pymysql
import re
import requests
import time

user = ""
cdb = ""
pasd = ""


tablename = "trans"
excelname = r".\transexample.xlsx"
crltitle = "chinese"

config = {"host" : "127.0.0.1",
          "port" : 3306,
          "user" : user,
          "password" : pasd,
          "db" : cdb,
          "charset" : "utf8mb4",
          "cursorclass" : pymysql.cursors.DictCursor}

headers = {"User-Agent":"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.146 Safari/537.36"}

# "User-Agent":"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"





def _output(ind,find = r"</p>\n <p>(.*)</p>\n   <p>以上为机器翻译结果，长、整句建议使用"):
    ''' find product data from get html data, ind = list of geted html data, front = from where to search key index '''
    report = ''
    pattern = re.compile(find)
    result = pattern.findall(ind)
    for i in result:
        report += i
    return report
    

def to_ydtrans(words, find = r"</p>\n <p>(.*)</p>\n   <p>以上为机器翻译结果，长、整句建议使用",delay = 3):
#     print("to_ydtrans : words -> %s"%(words))
    ''' put words to YouDao.com to translate, en = True means input words is english, this function translate between chinese and english '''

    url = ('http://dict.youdao.com/w/%s/#keyfrom=dict2.top'%(words))   
    time.sleep(delay)
    content = requests.get(url,headers = headers)
#     print("to_ydtrans : content.text -> %s"%(content.text))
    report = _output(content.text)
    print("to_ydtrans : %s done....."%(report))
    if report != []:
        return report
    else:
        return None

def create_table(sqlconfig,name,param):
    ''' create table to sqlconfig, name = tablename, argl = parameter of table '''
    inputparam = " (id INT UNSIGNED AUTO_INCREMENT, "+param+", PRIMARY KEY (id))"
    
    connection = pymysql.connect(**sqlconfig)
    cur = connection.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS " + name + inputparam)
    connection.commit()
    connection.close()
    print("create_table : \"" + tablename + "\" done.....")
    
def drop_table(sqlconfig,name):
    ''' drop table to sqlconfig, name = tablename '''
    connection = pymysql.connect(**sqlconfig)
    cur = connection.cursor()
    try:
        cur.execute("DROP TABLE " + name)
        connection.commit()
    except:
        print("Error NO table %s"%(name))
        connection.close()
        return

    connection.close()
    print("drop_table : \"" + tablename + "\" done.....")



def update_data(sqlconfig,sqls,tbname, dataname = '', where = '',row = 0):
    ''' update data to sqlconfig, sqls = input data list, tbname = table name, dataname = input data name, where = updata postion feature, row = data counts '''

    comd = "UPDATE " + tbname + " SET " + dataname + " = %s WHERE " + where + " = %s"
    i = 0
    count = row - 1
    while i < count:
        connection = pymysql.connect(**sqlconfig)
        cur = connection.cursor()
        cur.execute(comd,(sqls[i],str(i+1)))
        connection.commit()
        connection.close()
        i += 1
    print("update_data : \""+ dataname + " of " + tbname + "\" done.....")

       
def input_database(sqlconfig,sqls,tbname, dataname = '', datatype = ""):
    ''' input data to sqlconfig, sqls = input data list, tbname = table name, valuetype = input values type '''
    row = 1
    comd = "INSERT INTO " + tbname + " (" + dataname + ") VALUES (" + datatype + ")"
    
    for i in sqls:
        connection = pymysql.connect(**sqlconfig)
        cur = connection.cursor()
        cur.execute(comd,i)
        connection.commit()
        connection.close()
        row += 1
    print("input_database : \"" + tbname + "\" done.....")
    return row

def add_column(sqlconfig,tbname,dataname = '',datatype = ''):
    comd = "ALTER TABLE " + tbname + " ADD COLUMN " + dataname + " " + datatype
#     comd2 = "SELECT * FROM " + tbname + " LIMIT 1"
    
    comdL = [comd]
    
    for i in comdL:
        connection = pymysql.connect(**sqlconfig)
        cur = connection.cursor()
        cur.execute(i)
        connection.commit()
        connection.close()
    print("add_column : \"" + dataname + "\" to \"" + tablename + "\" done.....")
    
    
def output_data(sqlconfig,tablename,selectdata):
    ''' output data from sqlconfig, tablename = table name, selectdata = read data name '''
    resultlist = []
    connection = pymysql.connect(**sqlconfig)
    cur = connection.cursor()
    sql = "SELECT " + selectdata + " FROM " + tablename
    cur.execute(sql)
    result = cur.fetchall()  # get all data
    connection.commit()
    connection.close()
    
    for i in result:
        resultlist.append(i[selectdata]) 
    print("read_data : from \"" + tablename + "\" read data \"" + selectdata + "\" done.....")
    
    return resultlist

def load_excel(filename, crldatatitle):
    ''' load excel data column, filename = excel name and path, crldatatitle = chose column of excel '''
    workbook = xlrd.open_workbook(filename)
    sheetnamesL = []
    sheetsL = []
    
    allsheetsL = workbook.sheet_names()
    
    for i in allsheetsL:
        if not(i.startswith("Sheet")):
            sheetnamesL.append(i)
    
    for i in sheetnamesL:
        sheetsL.append(workbook.sheet_by_name(i))
    
    
    currentsheet = sheetsL[0]
    columnnum = currentsheet.ncols
    
    datatitleL = currentsheet.row_values(0)
    
    for i in range(columnnum):
        if crldatatitle == datatitleL[i]:
            takecol = currentsheet.col_values(i)
    
    #############   clean data
    resultL = []
    for i in takecol:
        result = re.sub('<([^>]+?)>','',i)
        result2 = re.sub('[a-zA-Z:/&;-]','',result)
        result3 = re.sub('[\d]','',result2)
        result4 = re.sub('[/n\._'   '' '"         "/|]','',result3)
        if result4 != '':
            resultL.append(result4)
    
#         resultL.append(i)
    
    return resultL



rel = load_excel(excelname, crltitle)

drop_table(config, tablename)

create_table(config,tablename, "chinese LONGTEXT NULL")

row = input_database(config,rel, tablename, "chinese", "%s")
# 
cp = output_data(config, tablename, "chinese")
# 
# 
nonecount = 0
storage = []
# count = 1
# all = 4
# 
for i in cp:
    result = to_ydtrans(i,False,delay = 5)
    if result == '':
        nonecount += 1
    storage.append(result)


#
add_column(config,tablename,"english","LONGTEXT NULL")
update_data(config,storage, tablename, "english", "id",row)
# 
# print("storage : %d  none : %d"%(len(storage),nonecount))


    
    
