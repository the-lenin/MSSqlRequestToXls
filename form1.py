import tkinter as tk
from tkinter import *
import pyodbc
import xlrd
import xlwt

class Application(tk.Frame):
    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.pack()
        self.createWidgets()

    def execute(self):
        self.connStr = 'DRIVER={SQL Server};SERVER='
        self.connStr += self.connStrIP.get() + ';DATABASE='
        self.connStr += self.connStrDbName.get() + ';UID=' + self.connStrUsr.get()
        self.connStr += ';PWD=' + self.connStrPass.get() + ';'

        try:
            fout = open('res.txt', 'w')
            cnxn = pyodbc.connect(self.connStr)
            cursor = cnxn.cursor()
            cursor.execute(self.connStrReq.get())
            self.rows = cursor.fetchall()

            for i in self.rows:
                fout.write(str(i) + '\n')
            print('Success!')
            for rr in cursor.description:
                print( rr)
            
        except Exception as Ex:
            print(Ex)
        finally:
            fout.close()

            
        try:
            font0 = xlwt.Font()
            font0.name = 'Arial 10'
            font0.colour_index = 0
            font0.bold = True

            style0 = xlwt.XFStyle()
            style0.font = font0

            style1 = xlwt.XFStyle()
            style1.num_format_str = 'D-MMM-YY'

            self.wb = xlwt.Workbook()
            self.ws = self.wb.add_sheet('Result 1')
            itera = 0

            for rr in cursor.description:
                self.ws.write(0, itera, str(rr[0]), style0)
                itera += 1
            
            itera = 1
            for kk in self.rows:
                it2 = 0
                for jj in kk:
                    if str(jj) == 'None':
                       self.ws.write(itera, it2, 'NULL', style0)
                    else:
                        self.ws.write(itera, it2, str(jj), style0)
                    it2 += 1
                itera += 1
            self.wb.save('example.xls')
           
        except Exception as Ex:
            print(Ex)

        
        

    def createWidgets(self):
        
        self.LblIP = Label(self, text = 'DataBase IP:', font = 'Arial 10')
        self.LblIP.pack()
        self.connStrIP = Entry(self, width = 100, bd = 1)
        self.connStrIP.insert(0, '172.17.1.12')
        self.connStrIP.pack()

        self.LblDbName = Label(self, text = 'DataBaseName:', font = 'Arial 10')
        self.LblDbName.pack()
        self.connStrDbName = Entry(self, width = 100, bd = 1)
        self.connStrDbName.insert(0, 'KODINSK_N')
        self.connStrDbName.pack()

        self.LblUsr = Label(self, text = 'Login:', font = 'Arial 10')
        self.LblUsr.pack()
        self.connStrUsr = Entry(self, width = 100, bd = 1)
        self.connStrUsr.insert(0, 'db2admin')
        self.connStrUsr.pack()

        self.LblPass = Label(self, text = 'Passwd:', font = 'Arial 10')
        self.LblPass.pack()
        self.connStrPass = Entry(self, width = 100, bd = 1)
        self.connStrPass.insert(0, '123')
        self.connStrPass.pack()

        self.LblReq = Label(self, text = 'Request:', font = 'Arial 10')
        self.LblReq.pack()
        self.connStrReq = Entry(self, width = 100, bd = 1)
        self.connStrReq.insert(0, 'SELECT * FROM CliAccount;')
        self.connStrReq.pack()

        self.OK = Button(self, text="OK", fg="green", command = self.execute)
        self.OK.pack()
        
        self.QUIT = tk.Button(self, text="QUIT", fg="red", command=root.destroy)
        self.QUIT.pack(side = BOTTOM)





root = Tk()
root.geometry('950x850+300+50')
root.title('Request to file')
app = Application(master=root)
app.mainloop()




