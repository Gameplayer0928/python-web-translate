# coding=utf-8
'''
Created on 2018年3月6日
 @author: Gameplayer0928 Qi Gao

#
#    This file is part of Exceltrans with YouDao v0.2.
#
#    Exceltrans with YouDao v0.2  is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    Exceltrans with YouDao v0.2 is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with Exceltrans with YouDao v0.2.  If not, see <http://www.gnu.org/licenses/>.
#
#    Copyright 2018, 2019, 2020 Qi Gao
# 

'''

''' this program is translate chinese to english with YouDao web. 
    create Table in MySQL database,
    read Excel file and put data to MySQL database,
    read data from database an put to YouDao web for translate,
    add 'english' data type in database and put translated data into database in 'english'.
    
    v0.2 add gui 
 '''

import tkinter
import tkinter.filedialog
import tkinter.messagebox

import xlrd

import pymysql
import re
import requests
import time

class TitleInput():
    def __init__(self,titlename,setin):
        self.frameinput = tkinter.LabelFrame(setin)
        self.textname = tkinter.Label(self.frameinput,text = titlename + ': ')
        self.textname.pack(side='left')
        self.textnamein = tkinter.Entry(self.frameinput)
        self.textnamein.pack(side="right")
        self.frameinput.pack()
    
    def get_data(self):
        rp = self.textnamein.get()
#         self.textnamein.clipboard_clear()
        return rp

class Exceltrans():
    def __init__(self):
#         self.user = ""
#         self.cdb = ""
#         self.pasd = ""


        self.tablename = ""
        self.excelname = ""
        self.crltitle = "chinese"

        self.config = {"host" : "127.0.0.1",
                       "port" : 3306,
                       "user" : '',
                       "password" : '',
                       "db" : '',
                       "charset" : "utf8mb4",
                       "cursorclass" : pymysql.cursors.DictCursor}
        
        self.headers = {"User-Agent":"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.146 Safari/537.36"}
        

        
        self.maingui = tkinter.Tk()
        self.maingui.title("Exceltrans with YouDao v0.2 - Gameplayer0928")
        
        self.filepath = tkinter.StringVar(self.maingui)
        self.filepath.set("no path")
        
        
        self.frame = tkinter.LabelFrame(self.maingui)
        
        self.filepathlabel = tkinter.Label(self.frame, textvariable = self.filepath)
        self.button = tkinter.Button(self.frame,text='select file to translate',command = self.load_file)

        self.button.pack()
        self.filepathlabel.pack()
        
        
        self.un = TitleInput("Mysql username",self.frame)
        self.up = TitleInput("Mysql password",self.frame)
        self.db = TitleInput("using database",self.frame)
        self.dt = TitleInput("data table",self.frame)
        
        self.frame.pack()
        self.configgetbutton = tkinter.Button(self.frame,text='set config',command = self.get_cfg)
        self.configgetbutton.pack()
        
        self.startbutton = tkinter.Button(self.frame,text='start translate',command = self.start)
        self.startbutton.pack()
        
        
        self.frame.pack()
         
        

    def load_file(self):
        ''' set excel file path '''
        fl = tkinter.filedialog.FileDialog(self.maingui)
        cd = fl.go('./')
        tkinter.messagebox.showinfo("file set", cd)
        self.excelname = cd
        self.filepath.set(cd)

    def get_cfg(self):
        self.config['user'] = self.un.get_data()
        self.config['password'] = self.up.get_data()
        self.config['db'] = self.db.get_data()
        self.tablename = self.dt.get_data()
        tkinter.messagebox.showinfo("config set", "config has been set")
        

    def show(self):
        ''' show program gui '''
        self.maingui.mainloop()

    def start(self):
        ''' start all process '''

#         print(self.config)
        self.drop_table(self.config, self.tablename)

        rel = self.load_excel(self.excelname, self.crltitle)
        
        self.create_table(self.config,self.tablename, "chinese LONGTEXT NULL")

        row = self.input_database(self.config,rel, self.tablename, "chinese", "%s")

        cp = self.output_data(self.config, self.tablename, "chinese")
      
        storage = self.all_toyoudao(cp)

        self.add_column(self.config,self.tablename,"english","LONGTEXT NULL")
        self.update_data(self.config,storage, self.tablename, "english", "id",row)



    def _output(self,ind,find = r"</p>\n <p>(.*)</p>\n   <p>以上为机器翻译结果，长、整句建议使用"):
        ''' find product data from get html data, ind = list of geted html data, front = from where to search key index '''
        report = ''
        pattern = re.compile(find)
        result = pattern.findall(ind)
        for i in result:
            report += i
        return report
        
    
    def to_ydtrans(self,words, find = r"</p>\n <p>(.*)</p>\n   <p>以上为机器翻译结果，长、整句建议使用",delay = 3):
    #     print("to_ydtrans : words -> %s"%(words))
        ''' put words to YouDao.com to translate, en = True means input words is english, this function translate between chinese and english '''
    
        url = ('http://dict.youdao.com/w/%s/#keyfrom=dict2.top'%(words))   
        time.sleep(delay)
        content = requests.get(url,headers = self.headers)
    #     print("to_ydtrans : content.text -> %s"%(content.text))
        report = self._output(content.text)
        print("to_ydtrans : %s done....."%(report))
        if report != []:
            return report
        else:
            return None
    
    def create_table(self,sqlconfig,name,param):
        ''' create table to sqlconfig, name = tablename, argl = parameter of table '''
        inputparam = " (id INT UNSIGNED AUTO_INCREMENT, "+param+", PRIMARY KEY (id))"
        
        connection = pymysql.connect(**sqlconfig)
        cur = connection.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS " + name + inputparam)
        connection.commit()
        connection.close()
        print("create_table : \"" + self.tablename + "\" done.....")
        
    def drop_table(self,sqlconfig,name):
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
        print("drop_table : \"" + self.tablename + "\" done.....")
    
    
    
    def update_data(self,sqlconfig,sqls,tbname, dataname = '', where = '',row = 0):
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
    
           
    def input_database(self,sqlconfig,sqls,tbname, dataname = '', datatype = ""):
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
    
    def add_column(self,sqlconfig,tbname,dataname = '',datatype = ''):
        ''' add a column in sqlconfin, tbname = table name, dataname = add data name, datatype = type of add data'''
        comd = "ALTER TABLE " + tbname + " ADD COLUMN " + dataname + " " + datatype
    #     comd2 = "SELECT * FROM " + tbname + " LIMIT 1"
        
        comdL = [comd]
        
        for i in comdL:
            connection = pymysql.connect(**sqlconfig)
            cur = connection.cursor()
            cur.execute(i)
            connection.commit()
            connection.close()
        print("add_column : \"" + dataname + "\" to \"" + self.tablename + "\" done.....")
        
        
    def output_data(self,sqlconfig,tablename,selectdata):
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
    
    def load_excel(self,filename, crldatatitle):
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
        print("load_excel : done.....")
        return resultL

    def all_toyoudao(self,cp):
        ''' translate one by one '''
        storage = []
        for i in cp:
            result = self.to_ydtrans(i,False,delay = 3)
            storage.append(result)
        print("all_toyoudao : done.....")
        return storage



####  start
if __name__ == "__main__":
    gogogo = Exceltrans()
    gogogo.show()
    
