# -*- coding: UTF-8 -*-
# Code by Pegasis
# Lua风格代码

import tkinter.messagebox  # 提示框库
import urllib.request  # 网络库
from _thread import start_new_thread  # 线程库
from os import _exit  # 退出
from os import remove  # 删除
from os.path import isfile  # 是否存在文件
from shutil import copy  # 复制
from tkinter import *  # UI库
from tkinter import font  # 字体库
from tkinter import ttk  # 进度条库
from tkinter.filedialog import asksaveasfilename, askopenfilename  # 选择文件库

import matplotlib.pyplot  as     plt  # 折线图库
import xlrd  # 电子表格库
from PyPDF2 import PdfFileReader, PdfFileWriter  # PDF库
from pylab import mpl  # 折线图辅助库
from pylab import xticks, yticks, ylim  # 折线图辅助库
from win32api import ShellExecute  # 打开文件库


# 合并pdf
def mergePdf(infnList, outPath):
    pdf_output = PdfFileWriter()
    pdfList = []
    for i in range(0, len(infnList)):
        infn = infnList[i]
        pdfList.append(open(infn, "rb"))
        pdf_input = PdfFileReader(pdfList[i])
        pdf_output.addPage(pdf_input.getPage(0))
    pdf_output.write(open(outPath, "wb"))
    for i in range(0, len(infnList)):
        pdfList[i].close()


def readFile(path):
    global students
    global studentnum
    global examNames
    data = xlrd.open_workbook(path)
    table = data.sheets()[0]
    studentnum = table.cell(0, 0).value
    print(studentnum)

    students = []
    for i in range(1, table.nrows):
        students.append(table.row_values(i))
        print(table.row_values(i))

    examNames = table.row_values(0)
    examNames = examNames[1:]
    print(examNames)


# 绘图
mpl.rcParams["font.sans-serif"] = ["SimHei"]  # 指定默认字体
mpl.rcParams["axes.unicode_minus"] = False


def drawChart(a, b):
    global students
    global studentnum
    global examNames
    global magicing
    global thenExit
    readFile(filePath)
    addDone()

    x = []
    for i in range(1, len(examNames) + 1):
        x.append(i)
    # 计算行数
    if len(students) // 2 == len(students) / 2:
        plty = len(students) // 2
    else:
        plty = len(students) // 2 + 1
    # 页数
    if plty // 3 == plty:
        page = plty // 3
    else:
        page = plty // 3 + 1
    # 计算Y轴刻度
    ytick = [1]
    if studentnum // 100 == studentnum:
        for i in range(100, int(studentnum + 1), 100):
            ytick.append(i)
    else:
        for i in range(100, int(studentnum), 100):
            ytick.append(i)
    # 循环
    pdfFiles = []
    for p in range(1, page + 1):  # 页循环
        plt.subplots_adjust(hspace=0.4)  # 图表纵向间隔
        for s in range(0, 6):  # 子图表循环
            if s + (p - 1) * 6 + 1 <= len(students):  # 是否有数据
                a = plt.subplot(3, 2, s + 1)  # 纵向,横向,位置
                plt.plot(x, students[s + (p - 1) * 6][1:], "black", linewidth=2)  # 载入参数
                xticks(x, examNames)  # X轴文字
                yticks(ytick, ytick)  # Y轴文字
                ylim(1, studentnum)  # 固定Y轴范围
                plt.gca().invert_yaxis()  # 反转Y轴方向
                plt.title(students[s + (p - 1) * 6][0])  # 图表标题
                ax = plt.gca()
                ax.spines["right"].set_color('none')
                ax.spines["top"].set_color('none')
                print("Chart" + str(p) + "," + str(s))
                addDone()

        fig = plt.gcf()
        fig.set_size_inches(8.27, 11.69)
        fig.savefig("Chart" + str(p) + ".pdf")
        pdfFiles.append("Chart" + str(p) + ".pdf")
        plt.close("all")
        if thenExit == 1:
            for files in pdfFiles:
                remove(files)
            _exit(0)
    mergePdf(pdfFiles, "Temp_Chart.pdf")

    for i in range(1, page + 1):
        remove("Chart" + str(i) + ".pdf")
    addDone()

    magicing = 0
    nextbtn.place(x=500, y=210, anchor="center")
    doingMagic["text"] = "完成！"


# 创建窗口
win = Tk()
win.title("Chart Maker")
# win.iconbitmap("icon.ico")
win.geometry("600x250")
win.resizable(0, 0)
thenExit = 0
V = "V1.4"


def closefunc():
    global magicing
    global thenExit
    if magicing == 0:
        if isfile("Temp_Chart.pdf"):
            remove("Temp_Chart.pdf")
        _exit(0)
    else:
        doingMagic["text"] = "准备退出....."
        thenExit = 1


win.protocol("WM_DELETE_WINDOW", closefunc)

# 步骤显示器
stepShower = Canvas(win, width=600, height=100)
# step shower polygons
SSPGs = [
    [  # 一个箭头
        [60, 30, 200, 30, 220, 50, 200, 70, 60, 70],  # 多边形坐标
        None,  # 多边形对象ID
        [135, 50],  # 文字坐标
        "导入数据",  # 文字
    ],
    [
        [210, 30, 350, 30, 370, 50, 350, 70, 210, 70, 230, 50],
        None,
        [285, 50],
        "制作图表",
    ],
    [
        [360, 30, 500, 30, 520, 50, 500, 70, 360, 70, 380, 50],
        None,
        [440, 50],
        "保存/打印",
    ],
]
for i in range(0, 3):
    SSPGs[i][1] = stepShower.create_polygon(*SSPGs[i][0], fill="", outline="black")
    stepShower.create_text(*SSPGs[i][2], text=SSPGs[i][3], font=("", 17))
stepShower.itemconfig(SSPGs[0][1], outline="#00dd00", width=3)
stepShower.place(x=0, y=0, anchor="nw")

# 下一步按钮
state = 1
filePath = ""
students = []
studentnum = 0
examNames = []
magicing = 0


def nextStep():
    global state
    global filePath
    global magicing
    if filePath != "" and magicing != 1:
        state = state + 1
        if state > 3:
            state = 1
        if state == 1:
            closefunc()
        elif state == 2:
            stepShower.itemconfig(SSPGs[1][1], outline="#00dd00", width=3)
            stepShower.itemconfig(SSPGs[0][1], outline="#000000", width=1)

            doingMagic.place(x=300, y=125, anchor="center")
            pbar.place(x=300, y=160, anchor="center")

            inputbtn.place_forget()
            helpText1.place_forget()
            helpText2.place_forget()
            fileNameText.place_forget()
            nextbtn.place_forget()

            magicing = 1
            start_new_thread(drawChart, (1, 1))
        else:
            stepShower.itemconfig(SSPGs[2][1], outline="#00dd00", width=3)
            stepShower.itemconfig(SSPGs[1][1], outline="#000000", width=1)

            seebtn.place(x=170, y=140, anchor="center")
            savebtn.place(x=300, y=140, anchor="center")
            printbtn.place(x=430, y=140, anchor="center")
            nextbtn["text"] = " 退出 "

            doingMagic.place_forget()
            pbar.place_forget()


nextbtn = ttk.Button(win, text="下一步", state="disabled", command=nextStep)
nextbtn.place(x=500, y=210, anchor="center")


# 第一步控件
def chooseFile():
    global filePath
    filePath = askopenfilename(filetypes=(("Microsoft Excel", "*.xls;*.xlsx"), ("All Files", ".*")))
    if filePath != "":
        fileNameText["text"] = filePath
        nextbtn["state"] = "normal"
        print(filePath)


ttk.Style().configure("size17.TButton", font=("", "17"))
inputbtn = ttk.Button(win, text="选择文件", style="size17.TButton", command=chooseFile)
inputbtn.place(x=300, y=140, anchor="center")

helpText1 = Label(win, text="第一次使用？", font=("", 10))
helpText1.place(x=310, y=100, anchor="ne")


def helpButton(a):
    ShellExecute(0, "open", "help.pdf", "", "", 1)


underlineFont = font.Font(family="Fixdsys", size=10, weight=font.NORMAL, underline=1)
helpText2 = Label(win, text="查看帮助", font=underlineFont)
helpText2.grid(row=10, column=0, columnspan=8)
helpText2.bind("<ButtonPress-1>", helpButton)
helpText2.place(x=310, y=100, anchor="nw")

fileNameText = Label(win, text="", fg="blue")
fileNameText.place(x=300, y=170, anchor="center")

# 第二步控件
doingMagic = Label(win, text="请稍后", font=("", 15))
doneNum = 0


def addDone():
    global doneNum
    allNum = len(students) + 2
    doneNum = doneNum + 1
    progressVar.set(doneNum / allNum * 100)


progressVar = DoubleVar()
pbar = ttk.Progressbar(win, orient=HORIZONTAL, length=300, mode="determinate", variable=progressVar)

# 第三步控件
ttk.Style().configure("size15.TButton", font=("", "15"))


def openPdf():
    ShellExecute(0, "open", "Temp_Chart.pdf", "", "", 1)


seebtn = ttk.Button(win, text="浏览", style="size15.TButton", command=openPdf)


def savePdf():
    path = asksaveasfilename(initialfile="分数图表.pdf", filetypes=(("PDF", "*.pdf"), ("All Files", ".*")))
    if path != "":
        copy("Temp_Chart.pdf", path)


savebtn = ttk.Button(win, text="保存", style="size15.TButton", command=savePdf)


def printPdf():
    try:
        ShellExecute(0, "print", "Temp_Chart.pdf", "", "", 1)
    except:
        tkinter.messagebox.showinfo("提示", "    未连接至打印机")


printbtn = ttk.Button(win, text="打印", style="size15.TButton", command=printPdf)


# 关于
def aboutButton(a):
    tkinter.messagebox.showinfo("关于",
                                "    Chart Maker " + V + "\n    By Pegasis\n    作者QQ 2766141475\n\n\n    ©Pegasis 2018   All Rights Reserved")


about = Label(win, text="关于", font=("", 10))
about.grid(row=10, column=0, columnspan=8)
about.bind("<ButtonPress-1>", aboutButton)
about.place(x=1, y=249, anchor="sw")


# 检查新版本
def newVersion(a):
    ShellExecute(0, "open", "https://pan.baidu.com/s/1nxl6QuP", "", "", 1)


checkNewVersionl = Label(win, text="有新版本", font=("", 10))
checkNewVersionl.bind("<ButtonPress-1>", newVersion)


# 下载新版本号
def makeVersionList(s):
    returnList = []
    i = 1
    while i <= len(s) / 2:
        i = i + 1
        returnList.append(int(s[i * 2 - 3]))
    return returnList


def cprsVersion(now, latest):
    nowl = makeVersionList(now)
    latestl = makeVersionList(latest)
    lgap = len(nowl) - len(latestl)
    if lgap > 0:
        for i in range(0, lgap):
            latestl.append(0)
    elif lgap < 0:
        for i in range(0, -lgap):
            nowl.append(0)

    for i in range(0, len(nowl)):
        if latestl[i] > nowl[i]:
            return True


def checkNewVersion(a, b):
    url = "https://raw.githubusercontent.com/Pegasis2003/ChartMakerUpdate/master/LatestVersion.txt"
    urllib.request.urlretrieve(url, "LatestVersion.txt")
    latestVersion = open("LatestVersion.txt", "r").readline()
    latestVersion = latestVersion[:-1]
    if cprsVersion(V, latestVersion):
        checkNewVersionl.place(x=40, y=249, anchor="sw")


start_new_thread(checkNewVersion, (1, 1))

win.mainloop()
