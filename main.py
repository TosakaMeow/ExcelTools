import webbrowser
from tkinter import *
import tkinter.messagebox
import win32ui
import win32con
import xlrd
import xlwt
import functions

file_type = 'Excel文件(*.xls .xlsx)|*.xls;*.xlsx|' \
            'All File(*.*)|*.*|' \
            '|'

API_flag = win32con.OFN_OVERWRITEPROMPT | win32con.OFN_FILEMUSTEXIST


def Win_Open_File_0():
    global file_0
    dlg = win32ui.CreateFileDialog(1, None, None, API_flag, file_type)  # 指定为打开文件窗口
    dlg.SetOFNInitialDir("C:")
    dlg.DoModal()
    path = dlg.GetPathName()
    file_0 = str(path)
    fileName_0.insert('end', file_0)
    print(file_0)


def Win_Open_File_1():
    global file_1
    dlg = win32ui.CreateFileDialog(1, None, None, API_flag, file_type)  # 指定为打开文件窗口
    dlg.SetOFNInitialDir("C:")
    dlg.DoModal()
    path = dlg.GetPathName()
    file_1 = str(path)
    fileName_1.insert('end', file_1)
    print(file_1)


def Load_Excel():
    try:
        data0 = xlrd.open_workbook(file_0)
        data1 = xlrd.open_workbook(file_1)
        print(data0.sheet_names())
        print(data1.sheet_names())
        functions.MainSolution(data0, data1)
    except:
        tkinter.messagebox.showerror('加载失败', '请检查文件是否被其他程序占用或未选取文件！')


def test(a):
    sheet = a.sheet_names()
    print(sheet)


def opensource():
    webbrowser.open("https://github.com/VioletLing/ExcelTools")


def intro():
    introduction = Tk()
    introduction.title("软件说明")
    introduction.geometry("430x200")
    Label(introduction, text="本软件用于合并两个Excel表,用法如下：").pack()
    Label(introduction, text="1.点击#选择第一个文件#, 在弹出的窗口中选择并确定").pack()
    Label(introduction, text="2.点击#选择第二个文件#, 在弹出的窗口中选择并确定").pack()
    Label(introduction, text="3.点击#加载文件#，如果没有出错则会加载成功并进入处理页面").pack()
    Label(introduction, text="4.前面会大致输出两个文件的部分信息，默认处理第一个Sheet").pack()
    Label(introduction, text="5.选择对应的基准组，第一个必选，第二个可选").pack()
    Label(introduction, text="6.确定后点击#开始合并文件#，合并后的文件将保存在软件目录下").pack()
    Label(introduction, text="其他问题请联系作者Violetnris@outlook.com").pack()


MainWindows = Tk()  # 主窗体
MainWindows.title("合并Excel小工具 Beta V1.0")
MainWindows.geometry("500x300")
menubar = Menu(MainWindows)
MainWindows.config(menu=menubar)
menubar.add_command(label="查看源代码", command=lambda: opensource())
menubar.add_command(label="说明", command=lambda: intro())
Button(text='选择第一个文件', command=lambda: Win_Open_File_0()).place(x=390, y=40)
fileName_0 = Text(MainWindows, width=50, height=2)
fileName_0.place(x=10, y=40)
Button(text='选择第二个文件', command=lambda: Win_Open_File_1()).place(x=390, y=70)
fileName_1 = Text(MainWindows, width=50, height=2)
fileName_1.place(x=10, y=70)
Button(text='加载文件', command=lambda: Load_Excel()).place(x=240, y=120)

MainWindows.mainloop()
