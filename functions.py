from tkinter import ttk

import xlrd
import xlwt
from tkinter import *
import win32ui
import win32con

file_type = 'Excel文件(*.xls)|*.xls|' \
            'All File(*.*)|*.*|' \
            '|'

API_flag = win32con.OFN_OVERWRITEPROMPT | win32con.OFN_FILEMUSTEXIST


def Win_Save_File():  # 保存文件时，文件后缀需要另处理
    print("Save File\n")
    dlg = win32ui.CreateFileDialog(0, None, None, API_flag, file_type)  # 指定为保存文件窗口
    dlg.SetOFNInitialDir('C:')  # 默认打开的位置
    dlg.DoModal()
    path = dlg.GetPathName()  # 获取打开的路径
    return path


def merge(file_0, file_1, cow0_0, cow0_1, cow1_0, cow1_1, display):
    print("正在合并文件！")

    fileName = Win_Save_File()
    if len(str(fileName)) == 0:
        return 0
    fileName = str(fileName) + ".xls"

    workbook = xlwt.Workbook(encoding='ascii')
    worksheet = workbook.add_sheet('My Worksheet')

    file_0_len = file_0.nrows
    file_1_len = file_1.nrows
    file_0_wid = file_0.ncols
    file_1_wid = file_1.ncols
    sum = int(0)
    if cow1_0 == "无":
        for i in range(file_0_len):
            master = file_0.cell(i, int(cow0_0) - 1).value
            for j in range(file_1_len):
                slave = file_1.cell(j, int(cow0_1) - 1).value
                if master == slave:
                    for n in range(file_0_wid):
                        worksheet.write(sum, n, label=file_0.cell(i, n).value)
                    for m in range(file_1_wid):
                        worksheet.write(sum, int(file_0_wid) + int(m), label=file_1.cell(i, m).value)
                    display.insert('end', master)
                    display.insert('end', "\n")
                    sum = sum + 1
                    break
    else:
        for i in range(file_0_len):
            master0 = file_0.cell(i, int(cow0_0) - 1).value
            master1 = file_0.cell(i, int(cow1_0) - 1).value
            for j in range(file_1_len):
                slave0 = file_1.cell(j, int(cow0_1) - 1).value
                slave1 = file_1.cell(j, int(cow1_1) - 1).value
                if master0 == slave0 and master1 == slave1:
                    for n in range(file_0_wid):
                        worksheet.write(sum, n, label=file_0.cell(i, n).value)
                    for m in range(file_1_wid):
                        worksheet.write(sum, int(file_0_wid) + int(m), label=file_1.cell(i, m).value)
                        display.insert('end', master0)
                        display.insert('end', "\n")
                    sum = sum + 1
                    break

    workbook.save(fileName)
    display.insert('end', "文件合并成功，保存在")
    display.insert('end', fileName)


def MainSolution(file_0, file_1):
    window = Tk()
    window.title("加载完毕...")
    window.geometry("1000x600")
    file_0_sheet = file_0.sheet_names()
    file_1_sheet = file_1.sheet_names()
    Label(window, text="文件基本信息").place(x=465, y=1)

    Label(window, text="第一个文件需要加载的Sheet").place(x=40, y=20)
    Label(window, text=file_0_sheet[0]).place(x=200, y=20)
    Label(window, text="第二个文件需要加载的Sheet").place(x=40, y=40)
    Label(window, text=file_1_sheet[0]).place(x=200, y=40)

    sheet_0 = file_0.sheet_by_name(str(file_0_sheet[0]))
    sheet_1 = file_1.sheet_by_name(str(file_1_sheet[0]))

    Label(window, text="第一个表的总行数").place(x=40, y=60)
    Label(window, text=sheet_0.nrows).place(x=200, y=60)

    Label(window, text="第一个表的总列数").place(x=240, y=60)
    Label(window, text=sheet_0.ncols).place(x=400, y=60)

    Label(window, text="第二个表的总行数").place(x=40, y=80)
    Label(window, text=sheet_1.nrows).place(x=200, y=80)

    Label(window, text="第二个表的总列数").place(x=240, y=80)
    Label(window, text=sheet_1.ncols).place(x=400, y=80)

    Label(window, text="合并文件设置").place(x=465, y=120)

    Label(window, text="选取第一个文件的第一个基准组").place(x=40, y=160)
    xVariable0 = StringVar()  # #创建变量，便于取值
    FF0 = ttk.Combobox(window, textvariable=xVariable0)  # #创建下拉菜单
    FF0.place(x=240, y=160)  # #将下拉菜单绑定到窗体
    FF0["value"] = ("1", "2", "3", "4", "5", "6", "7", "8", "9")  # #给下拉菜单设定值
    FF0.current(0)

    SFT = Label(window, text="选取第二个文件的第一个基准组")
    SFT.place(x=480, y=160)
    xVariable1 = StringVar()  # #创建变量，便于取值
    FF1 = ttk.Combobox(window, textvariable=xVariable1)  # #创建下拉菜单
    FF1.place(x=680, y=160)  # #将下拉菜单绑定到窗体
    FF1["value"] = ("1", "2", "3", "4", "5", "6", "7", "8", "9")  # #给下拉菜单设定值
    FF1.current(0)

    FST = Label(window, text="选取第一个文件的第二个基准组")
    FST.place(x=40, y=180)
    xVariable2 = StringVar()  # #创建变量，便于取值
    FS0 = ttk.Combobox(window, textvariable=xVariable2)  # #创建下拉菜单
    FS0.place(x=240, y=180)  # #将下拉菜单绑定到窗体
    FS0["value"] = ("无", "1", "2", "3", "4", "5", "6", "7", "8", "9")  # #给下拉菜单设定值
    FS0.current(0)

    FST = Label(window, text="选取第二个文件的第二个基准组")
    FST.place(x=480, y=180)
    xVariable3 = StringVar()  # #创建变量，便于取值
    FS1 = ttk.Combobox(window, textvariable=xVariable3)  # #创建下拉菜单
    FS1.place(x=680, y=180)  # #将下拉菜单绑定到窗体
    FS1["value"] = ("无", "1", "2", "3", "4", "5", "6", "7", "8", "9")  # #给下拉菜单设定值
    FS1.current(0)

    SM = Button(window, text='开始合并文件',
                command=lambda: merge(sheet_0, sheet_1, FF0.get(), FF1.get(), FS0.get(), FS1.get(), message_display))
    SM.place(x=80, y=220)

    Label(window, text="合并日志").place(x=270, y=250)

    message_display = Text(window)
    message_display.place(x=40, y=275)

    window.mainloop()
