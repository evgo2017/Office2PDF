"""
【程序功能】：将文件夹内所有的 ppt、excel、word 均生成一份对应的 PDF 文件
【作者】：evgo（evgo2017.com）
【目标文件夹】：默认为此程序目前所在的文件夹；
                若输入路径，则为该文件夹（只转换该层，不转换子文件夹下内容）
【生成的pdf名称】：原始名称+.pdf
"""
import os, win32com.client, gc, tkinter as tk
from tkinter import filedialog, IntVar, ttk
from enum import Enum

# TODO：解决转化完成才开始写入，子线程更新 UI

# 界面基础
window = tk.Tk()
windowHeight = 500
windowWidth = 600
fromFolderEntry = None
toFolderEntry = None
logListText = None
wordCheckVar = tk.IntVar()
pptCheckVar= tk.IntVar()
excelCheckVar = tk.IntVar()
allTypeCheckVar = tk.IntVar()

def chooseFolderPath():
    return tk.filedialog.askdirectory(initialdir=os.getcwd(), title="Select file")
def setFromFolderPath():
    folderPath = chooseFolderPath()
    if len(folderPath) > 0:
        fromFolderPath = os.path.normpath(folderPath)
        fromFolderEntry.delete(0, tk.END)
        fromFolderEntry.insert(0, fromFolderPath)
def setToFolderPath():
    folderPath = chooseFolderPath()
    if len(folderPath) > 0:
        toFolderPath = os.path.normpath(folderPath)
        toFolderEntry.delete(0, tk.END)
        toFolderEntry.insert(0, toFolderPath)
def getFromFolderPath():
    return fromFolderEntry.get()
def getToFolderPath():
    return toFolderEntry.get()
def toggleSelectAllConvertType():
    flag = allTypeCheckVar.get() == 1
    wordCheckVar.set(flag)
    pptCheckVar.set(flag)
    excelCheckVar.set(flag)
def setAllTypeCheckVar():
    allTypeCheckVar.set(wordCheckVar.get() + pptCheckVar.get() + excelCheckVar.get() == 3)

def insertLog(log):
    return logListText.insert(tk.END, log + "\n")
# 修改后缀名
def changeSufix2Pdf(file):
    return file[:file.rfind('.')]+".pdf"
# 添加工作簿序号
def addWorksheetsOrder(file, i):
    return file[:file.rfind('.')]+"_工作表"+str(i)+".pdf"
# 转换地址
def toFileJoin(filePath,file):
    return os.path.join(filePath, file[:file.rfind('.')]+".pdf")

# Word
def word2Pdf(fromFolderPath, toFolderPath, words):
    # 如果没有文件则提示后直接退出
    if(len(words)<1):
        insertLog("\n【无 Word 文件】\n")
        return
    # 开始转换
    insertLog("\n【开始 Word -> PDF 转换】")
    try:
        insertLog("打开 Word 进程...")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = 0
        word.DisplayAlerts = False
        doc = None
        for i in range(len(words)):
            insertLog("\n" + str(i))
            fileName = words[i] # 文件名称
            fromFilePath = os.path.join(fromFolderPath, fileName) # 文件地址
            toFileName = changeSufix2Pdf(fileName) # 生成的文件名称
            toFilePath = toFileJoin(toFolderPath, toFileName) # 生成的文件地址
            insertLog("转换："+ fileName +"文件中...")
            insertLog("文件来源地址：" + fromFilePath)
            insertLog("文件目标地址：" + toFilePath)
            # 某文件出错不影响其他文件打印
            try:
                doc = word.Documents.Open(fromFilePath)
                doc.SaveAs(toFilePath, 17) # 生成的所有 PDF 都会在 PDF 文件夹中
                insertLog("完成："+ toFileName)
            except Exception as e:
                insertLog(str(e))
            # 关闭 Word 进程
        insertLog("所有 Word 文件已打印完毕")
        insertLog("结束 Word 进程...\n")
        doc.Close()
        doc = None
        word.Quit()
        word = None 
    except Exception as e:
        insertLog(str(e))
    finally:
        gc.collect()

# Excel
def excel2Pdf(fromFolderPath, toFolderPath, excels):
    # 如果没有文件则提示后直接退出
    if(len(excels)<1):
        insertLog("\n【无 Excel 文件】\n")
        return
    # 开始转换
    insertLog("\n【开始 Excel -> PDF 转换】")
    try:
        insertLog("打开 Excel 进程中...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = 0
        excel.DisplayAlerts = False
        wb = None
        ws = None
        for i in range(len(excels)):
            insertLog(str(i))
            fileName = excels[i] # 文件名称
            fromFilePath = os.path.join(fromFolderPath, fileName) # 文件地址
            
            insertLog("转换："+fileName+"文件中...")
            # 某文件出错不影响其他文件打印
            try:
                wb = excel.Workbooks.Open(fromFilePath)
                for j in range(wb.Worksheets.Count): # 工作表数量，一个工作簿可能有多张工作表
                    toFileName = addWorksheetsOrder(fileName, j+1) # 生成的文件名称
                    toFilePath = toFileJoin(toFolderPath, toFileName) # 生成的文件地址
                    
                    ws = wb.Worksheets(j+1) # 若为[0]则打包后会提示越界
                    ws.ExportAsFixedFormat(0, toFilePath) # 每一张都需要打印
                    insertLog("转换至："+toFileName+"文件完成")
            except Exception as e:
                insertLog(str(e))
        # 关闭 Excel 进程
        insertLog("所有 Excel 文件已打印完毕")
        insertLog("结束 Excel 进程中...\n")
        ws = None
        wb.Close()
        wb = None
        excel.Quit()
        excel = None
    except Exception as e:
        insertLog(str(e))
    finally: 
        gc.collect()

# PPT
def ppt2Pdf(fromFolderPath, toFolderPath, ppts):
    # 如果没有文件则提示后直接退出
    if(len(ppts)<1):
        insertLog("\n【无 PPT 文件】\n")
        return
    # 开始转换
    insertLog("\n【开始 PPT -> PDF 转换】")
    try:
        insertLog("打开 PowerPoint 进程中...")
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        ppt = None
        # 某文件出错不影响其他文件打印

        for i in range(len(ppts)):
            insertLog(str(i))
            fileName = ppts[i] # 文件名称
            fromFilePath = os.path.join(fromFolderPath, fileName) # 文件地址
            toFileName = changeSufix2Pdf(fileName) # 生成的文件名称
            toFilePath = toFileJoin(toFolderPath, toFileName) # 生成的文件地址

            insertLog("转换：" + fileName + "文件中...")
            try:
                ppt = powerpoint.Presentations.Open(fromFilePath, WithWindow=False)
                if ppt.Slides.Count>0:
                    ppt.SaveAs(toFilePath, 32) # 如果为空则会跳出提示框（暂时没有找到消除办法）
                    insertLog("转换至："+toFileName+"文件完成")
                else:
                    insertLog("（错误，发生意外：此文件为空，跳过此文件）")
            except Exception as e:
                insertLog(str(e))
        # 关闭 PPT 进程
        insertLog("所有 PPT 文件已打印完毕")
        insertLog("结束 PowerPoint 进程中...")
        ppt.Close()
        ppt = None
        powerpoint.Quit()
        powerpoint = None
    except Exception as e:
        insertLog(str(e))
    finally:
        gc.collect()

# 核心转换
def startConvert():
    # TODO：是否遍历子目录
    # 将目标文件夹所有文件归类，转换时只打开一个进程
    words = []
    ppts = []
    excels = []
    
    fromFolderPath = getFromFolderPath()
    for fn in os.listdir(fromFolderPath):
        if fn.endswith(('.doc', 'docx')):
            words.append(fn)
        if fn.endswith(('.ppt', 'pptx')):
            ppts.append(fn)
        if fn.endswith(('.xls', 'xlsx')):
            excels.append(fn)

    insertLog("====================开始转换====================")

    toFolderPath = getToFolderPath()
    if not os.path.exists(toFolderPath):
        os.makedirs(toFolderPath)

    if (wordCheckVar.get() == 1):
      word2Pdf(fromFolderPath, toFolderPath, words)
    if (pptCheckVar.get() == 1):
      ppt2Pdf(fromFolderPath, toFolderPath, ppts)
    if (excelCheckVar.get() == 1):
      excel2Pdf(fromFolderPath, toFolderPath, excels)

    insertLog("====================转换结束====================")

def initView():
    global window, windowHeight, windowWidth, fromFolderEntry, toFolderEntry, logListText, wordCheckVar, pptCheckVar, excelCheckVar, allTypeCheckVar
    # window
    window.title("Office2PDF")
    screenwidth = window.winfo_screenwidth()
    screenheight = window.winfo_screenheight()
    window.geometry('%dx%d+%d+%d' % (windowWidth, windowHeight, (screenwidth  - windowWidth) / 2, (screenheight - windowHeight) / 2))

    frame = tk.Frame(window)
    frame.pack(padx = 10, pady = 10)

    # frame
    infoLabelFrame = ttk.LabelFrame(frame, text = "基本信息")
    folerLabelFrame = ttk.LabelFrame(frame, text = "文件夹")
    configFrame = tk.Frame(frame)
    startFrame = tk.Frame(frame)
    logListFrame = tk.Frame(frame)

    infoLabelFrame.pack(fill = tk.X, expand = tk.YES, pady = 4)
    folerLabelFrame.pack(fill = tk.X, expand = tk.YES, pady = 4, ipady = 4)
    configFrame.pack(fill = tk.X)
    startFrame.pack(fill = tk.X, expand = tk.YES, pady = 4)
    logListFrame.pack(fill = tk.BOTH, expand = tk.YES, pady = 4)

    # infoLableFrame
    fnFrame = tk.Frame(infoLabelFrame)
    authorFrame = tk.Frame(infoLabelFrame)

    fnFrame.pack(fill = tk.X, expand = tk.YES, padx = 2)
    authorFrame.pack(fill = tk.X, expand = tk.YES, padx = 2)

    # fnFrame
    fnLabel = tk.Label(fnFrame, text = "将文件夹内的 Word、Excel 或 PPT 生成对应的 PDF 文件")

    fnLabel.pack(side = tk.LEFT)

    # authorFrame
    # TODO：复制功能
    authorLabel = tk.Label(authorFrame, text = "作者：evgo（evgo2017.com）")
    
    authorLabel.pack(side = tk.RIGHT)

    # folderLableFrame
    fromFolderFrame = tk.Frame(folerLabelFrame)
    toFolderFrame = tk.Frame(folerLabelFrame)
    
    fromFolderFrame.pack(fill = tk.X, padx = 3)
    toFolderFrame.pack(fill = tk.X, padx = 3)

    # fromFolderFrame
    fromFolderLabel = tk.Label(fromFolderFrame, text = "来源")
    fromFolderEntry = ttk.Entry(fromFolderFrame)
    fromFolderButton = ttk.Button(fromFolderFrame, text = "选择", command = setFromFolderPath)

    fromFolderLabel.pack(side = tk.LEFT)
    fromFolderEntry.pack(side = tk.LEFT, fill = tk.X, expand = tk.YES, padx = 6)
    fromFolderButton.pack(side = tk.LEFT)

    # toFolderFrame
    toFolderLabel = tk.Label(toFolderFrame, text = "目标")
    toFolderEntry = ttk.Entry(toFolderFrame)
    toFolderButton = ttk.Button(toFolderFrame, text = "选择", command = setToFolderPath)

    toFolderLabel.pack(side = tk.LEFT)
    toFolderEntry.pack(side = tk.LEFT, fill = tk.X, expand = tk.YES, padx = 6)
    toFolderButton.pack(side = tk.LEFT)

    # configFrame
    convertTypeLabelFrame = tk.LabelFrame(configFrame, text="转换类型")
    convertTypeLabelFrame.pack(side = tk.LEFT)
    choseLabelFrame = tk.LabelFrame(configFrame, text="子文件夹")
    # choseLabelFrame.pack(side = tk.LEFT)
    v = IntVar()
    FRUITS = [
        ('转化', 1),
        ('不转化', 2)
    ]
    for name, num in FRUITS:
        tk.Radiobutton(choseLabelFrame, text=name, variable=v, value=num).pack(side = tk.LEFT)

    wordCheckbutton = tk.Checkbutton(convertTypeLabelFrame, text = 'Word', variable = wordCheckVar, command = setAllTypeCheckVar)
    pptCheckbutton = tk.Checkbutton(convertTypeLabelFrame, text = 'PPT', variable = pptCheckVar, command = setAllTypeCheckVar)
    excelCheckbutton = tk.Checkbutton(convertTypeLabelFrame, text = 'Excel', variable = excelCheckVar, command = setAllTypeCheckVar)
    allTypeCheckbutton = tk.Checkbutton(convertTypeLabelFrame, text="全选/全不选", variable = allTypeCheckVar, command = toggleSelectAllConvertType)

    wordCheckbutton.pack(side = tk.LEFT)
    pptCheckbutton.pack(side = tk.LEFT)
    excelCheckbutton.pack(side = tk.LEFT)
    allTypeCheckbutton.pack(side = tk.LEFT)

    # startFrame
    startButton = ttk.Button(startFrame, text = '开始', command = startConvert)
    startButton.pack(side = tk.LEFT, fill = tk.X, expand = tk.YES, ipady = 1.5)

    # logListFrame
    logListText = tk.Text(logListFrame, height = 100)
    logListText.pack(fill = tk.Y, expand = tk.YES)

def init():
    global window

    initView()
    fromFolderEntry.insert(0, os.getcwd())
    toFolderEntry.insert(0, os.getcwd())
    wordCheckVar.set(1)
    pptCheckVar.set(1)
    excelCheckVar.set(1)
    allTypeCheckVar.set(1)

    # 开始程序
    insertLog("【程序功能】将目标路径下内所有的 ppt、excel、word 均生成一份对应的 PDF 文件，存在新生成的 pdf 文件夹中（需已经安装office，不包括子文件夹）")
    insertLog("【作者】：evgo，evgo2017.com，公众号（随风前行），Github（evgo2017）")
    insertLog("注意：若某 PPT 和 Excel 文件为空，则会出错跳过此文件。若转换 PPT 时间过长，请查看是否有报错窗口等待确认，暂时无法彻底解决 PPT 的窗口问题（为空错误已解决）。在关闭进程过程中，时间可能会较长，十秒左右，请耐心等待。")
    
    # 进入消息循环
    window.mainloop()

init()
