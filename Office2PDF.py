"""
【程序功能】：将文件夹内所有的 ppt、excel、word 均生成一份对应的 PDF 文件
【作者】：evgo（evgo2017.com）
【目标文件夹】：默认为此程序目前所在的文件夹；
                若输入路径，则为该文件夹（只转换该层，不转换子文件夹下内容）
【生成的pdf名称】：原始名称+.pdf
"""
import os
import win32com.client
import gc
import tkinter as tk
from tkinter import filedialog, StringVar, IntVar, ttk
import queue
import time
import sys
import threading

window = tk.Tk()
logQueue = queue.Queue()

fromRootFolderPathVar = tk.StringVar()
toRootFolderPathVar = tk.StringVar()
isCovertChildrenFolderVar = tk.IntVar()
isCovertWordVar = tk.IntVar()
isCovertPPTVar= tk.IntVar()
isCovertExcelVar = tk.IntVar()
isCovertAllTypeVar = tk.IntVar()

# 赋默认值
fromRootFolderPathVar.set(os.getcwd())
toRootFolderPathVar.set(os.getcwd())
isCovertWordVar.set(1)
isCovertPPTVar.set(1)
isCovertExcelVar.set(1)
isCovertAllTypeVar.set(1)
isCovertChildrenFolderVar.set(1)

def setAllTypeCheckVar():
    isCovertAllTypeVar.set(isCovertWordVar.get() + isCovertPPTVar.get() + isCovertExcelVar.get() == 3)
def formatPath(path):
    return os.path.normpath(path)
def insertLog(log):
    logQueue.put(log)
def startConvert():
    T = threading.Thread(target = convert, args = (fromRootFolderPathVar.get(), toRootFolderPathVar.get()))
    T.daemon = True
    T.start()
def changeSufix2Pdf(file):
    return file[:file.rfind('.')]+".pdf"
def addWorksheetsOrder(file, i):
    return file[:file.rfind('.')] + "_" + str(i) + file[file.rfind('.'):]

# Word
def word2Pdf(fromRootFolderPath, toRootFolderPath, words):
    # 如果没有文件则提示后直接退出
    if(len(words)<1):
        insertLog("\n【无 Word 文件】\n")
        return
    # 开始转换
    insertLog("\n【 Word -> PDF 转换】")
    fromRootFolderPath = formatPath(fromRootFolderPath)
    toRootFolderPath = formatPath(toRootFolderPath)
    try:
        insertLog("\n打开 Word 进程中...")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = 0
        word.DisplayAlerts = False
        doc = None
        for i in range(len(words)):
            insertLog("\n" + str(i))
            fromFilePath = formatPath(words[i])
            fromFileName = os.path.basename(fromFilePath)
            insertLog("原始文件：" + fromFilePath)
            subPath = fromFilePath[len(fromRootFolderPath) + 1 : len(fromFilePath) - len(fromFileName)]
            toSubFolderPath = os.path.join(toRootFolderPath, subPath)
            # 子文件夹创建
            if not os.path.exists(toSubFolderPath):
                os.makedirs(toSubFolderPath)
            toFileName = changeSufix2Pdf(fromFileName)
            toFilePath = os.path.join(toSubFolderPath, toFileName)
            # 某文件出错不影响其他文件打印
            try:
                doc = word.Documents.Open(fromFilePath)
                doc.SaveAs(toFilePath, 17) # 生成的所有 PDF 都会在 PDF 文件夹中
                insertLog("生成文件：" + toFilePath)
            except Exception as e:
                insertLog(str(e))
            # 关闭 Word 进程
        insertLog("\n所有 Word 文件已转换完毕\n")
        insertLog("结束 Word 进程中...")
        doc.Close()
        doc = None
        word.Quit()
        word = None
        insertLog("已结束 Word 进程\n")
    except Exception as e:
        insertLog(str(e))
    finally:
        gc.collect()

# Excel
def excel2Pdf(fromRootFolderPath, toRootFolderPath, excels):
    # 如果没有文件则提示后直接退出
    if(len(excels)<1):
        insertLog("\n【无 Excel 文件】\n")
        return
    # 开始转换
    insertLog("\n【 Excel -> PDF 转换】")
    fromRootFolderPath = formatPath(fromRootFolderPath)
    toRootFolderPath = formatPath(toRootFolderPath)
    try:
        insertLog("\n打开 Excel 进程中...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = 0
        excel.DisplayAlerts = False
        wb = None
        ws = None
        for i in range(len(excels)):
            insertLog("\n" + str(i))
            fromFilePath = formatPath(excels[i])
            fromFileName = os.path.basename(fromFilePath)
            insertLog("原始文件：" + fromFilePath)
            subPath = fromFilePath[len(fromRootFolderPath) + 1 : len(fromFilePath) - len(fromFileName)]
            toSubFolderPath = os.path.join(toRootFolderPath, subPath)
            # 子文件夹创建
            if not os.path.exists(toSubFolderPath):
                os.makedirs(toSubFolderPath)
            # 某文件出错不影响其他文件打印
            try:
                wb = excel.Workbooks.Open(fromFilePath)
                count = wb.Worksheets.Count
                insertLog("此 Excel 一共有" + str(count) + "张表：")
                for j in range(count): # 工作表数量，一个工作簿可能有多张工作表
                    insertLog("\n转换第" + str(j + 1) + "张表中...")
                    if (count == 1):
                        toFileName = changeSufix2Pdf(fromFileName) # 生成的文件名称
                    else:
                        toFileName = changeSufix2Pdf(addWorksheetsOrder(fromFileName, j + 1)) # 仅多张表时加序号
                    toFilePath = os.path.join(toSubFolderPath, toFileName) # 生成的文件地址
                    insertLog("生成文件：" + toFilePath)
                    ws = wb.Worksheets(j + 1) # 若为[0]则打包后会提示越界
                    ws.ExportAsFixedFormat(0, toFilePath) # 每一张都需要打印
            except Exception as e:
                insertLog(str(e))
        # 关闭 Excel 进程
        insertLog("\n所有 Excel 文件已转换完毕\n")
        insertLog("结束 Excel 进程中...")
        ws = None
        wb.Close()
        wb = None
        excel.Quit()
        excel = None
        insertLog("已结束 Excel 进程\n")
    except Exception as e:
        insertLog(str(e))
    finally: 
        gc.collect()

# PPT
def ppt2Pdf(fromRootFolderPath, toRootFolderPath, ppts):
    # 如果没有文件则提示后直接退出
    if(len(ppts)<1):
        insertLog("\n【无 PPT 文件】\n")
        return
    # 开始转换
    insertLog("\n【 PPT -> PDF 转换】")
    fromRootFolderPath = formatPath(fromRootFolderPath)
    toRootFolderPath = formatPath(toRootFolderPath)
    try:
        insertLog("\n打开 PowerPoint 进程中...")
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        ppt = None
        # 某文件出错不影响其他文件打印

        for i in range(len(ppts)):
            insertLog("\n" + str(i))
            fromFilePath = formatPath(ppts[i])
            fromFileName = os.path.basename(fromFilePath)
            insertLog("原始文件：" + fromFilePath)
            subPath = fromFilePath[len(fromRootFolderPath) + 1 : len(fromFilePath) - len(fromFileName)]
            toSubFolderPath = os.path.join(toRootFolderPath, subPath)
            # 子文件夹创建
            if not os.path.exists(toSubFolderPath):
                os.makedirs(toSubFolderPath)
            toFileName = changeSufix2Pdf(fromFileName)
            toFilePath = os.path.join(toSubFolderPath, toFileName) # 生成的文件地址
            try:
                ppt = powerpoint.Presentations.Open(fromFilePath, WithWindow=False)
                if ppt.Slides.Count > 0:
                    ppt.SaveAs(toFilePath, 32) # 如果为空则会跳出提示框（暂时没有找到消除办法）
                    insertLog("生成文件：" + toFilePath)
                else:
                    insertLog("生成文件：（错误，发生意外：此文件为空，跳过此文件）")
            except Exception as e:
                insertLog(str(e))
        # 关闭 PPT 进程
        insertLog("\n所有 PPT 文件已转换完毕\n")
        insertLog("结束 PowerPoint 进程中...")
        insertLog("若耗时较久（超过30秒），建议再前往任务管理器手动关闭该进程。")
        ppt.Close()
        ppt = None
        powerpoint.Quit()
        powerpoint = None
        insertLog("已结束 PowerPoint 进程\n")
    except Exception as e:
        insertLog(str(e))
    finally:
        gc.collect()

# 核心转换
def convert(fromRootFolderPath, toRootFolderPath):
    # TODO：是否遍历子目录
    # 将目标文件夹所有文件归类，转换时只打开一个进程
    words = []
    ppts = []
    excels = []

    for folderPath, dirs, fileNames in os.walk(fromRootFolderPath):
        for fileName in fileNames:
            fromFilePath = formatPath(os.path.join(folderPath, fileName))
            if fileName.endswith(('.doc', 'docx')):
                words.append(fromFilePath)
            elif fileName.endswith(('.ppt', 'pptx')):
                ppts.append(fromFilePath)
            elif fileName.endswith(('.xls', 'xlsx')):
                excels.append(fromFilePath)
        if isCovertChildrenFolderVar.get() == 0:
            break

    insertLog("\n====================开始转换====================")

    if (isCovertWordVar.get() == 1):
      word2Pdf(fromRootFolderPath, toRootFolderPath, words)
    if (isCovertExcelVar.get() == 1):
      excel2Pdf(fromRootFolderPath, toRootFolderPath, excels)
    if (isCovertPPTVar.get() == 1):
      ppt2Pdf(fromRootFolderPath, toRootFolderPath, ppts)

    insertLog("====================转换结束====================")

class GUI():
    def __init__(self, window, windowHeight = 530, windowWidth = 500):
        self.window = window
        self.windowHeight = windowHeight
        self.windowWidth = windowWidth
        
        # 数据有交互的变量
        self.logListText = None
        self.fromFolderEntry = None
        self.toFolderEntry = None

        self.initGui()

        # 启动 after 方法
        self.window.after(100, self.showLog)
        # 进入消息循环
        self.window.mainloop()

    def initGui(self):
        # window
        self.window.title("Office2PDF")

        windowWidth = self.windowWidth
        windowHeight = self.windowHeight
        screenwidth = self.window.winfo_screenwidth()
        screenheight = self.window.winfo_screenheight()
        
        window.geometry('%dx%d+%d+%d' % (windowWidth, windowHeight, (screenwidth  - windowWidth) / 2, (screenheight - windowHeight) / 2))

        frame = tk.Frame(window)
        frame.pack(fill = tk.BOTH, expand = tk.YES, padx = 10, pady = 10)

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
        tk.Label(fnFrame, text = "将文件夹内的 Word、Excel 或 PPT 生成对应的 PDF 文件。").grid(sticky = tk.W)
        tk.Label(fnFrame, text = "需已安装 Office 2007 及以上，或 Microsoft Save as PDF 加载项。").grid(sticky = tk.W)

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
        fromFolderEntry = ttk.Entry(fromFolderFrame, textvariable = fromRootFolderPathVar)
        self.fromFolderEntry = fromFolderEntry
        fromFolderButton = ttk.Button(fromFolderFrame, text = "选择", command = self.setFromFolderPath)

        fromFolderLabel.pack(side = tk.LEFT)
        fromFolderEntry.pack(side = tk.LEFT, fill = tk.X, expand = tk.YES, padx = 6)
        fromFolderButton.pack(side = tk.LEFT)

        # toFolderFrame
        toFolderLabel = tk.Label(toFolderFrame, text = "目标")
        toFolderEntry = ttk.Entry(toFolderFrame, textvariable = toRootFolderPathVar)
        self.toFolderEntry = toFolderEntry
        toFolderButton = ttk.Button(toFolderFrame, text = "选择", command = self.setToFolderPath)

        toFolderLabel.pack(side = tk.LEFT)
        toFolderEntry.pack(side = tk.LEFT, fill = tk.X, expand = tk.YES, padx = 6)
        toFolderButton.pack(side = tk.LEFT)

        # configFrame
        convertTypeLabelFrame = tk.LabelFrame(configFrame, text="转换类型")
        concertChildrenFolderLabelFrame = tk.LabelFrame(configFrame, text="子文件夹")
        convertTypeLabelFrame.pack(side = tk.LEFT)
        concertChildrenFolderLabelFrame.pack(side = tk.LEFT)

        wordCheckbutton = tk.Checkbutton(convertTypeLabelFrame, text = 'Word', variable = isCovertWordVar, command = setAllTypeCheckVar)
        pptCheckbutton = tk.Checkbutton(convertTypeLabelFrame, text = 'PPT', variable = isCovertPPTVar, command = setAllTypeCheckVar)
        excelCheckbutton = tk.Checkbutton(convertTypeLabelFrame, text = 'Excel', variable = isCovertExcelVar, command = setAllTypeCheckVar)
        allTypeCheckbutton = tk.Checkbutton(convertTypeLabelFrame, text="全选/全不选", variable = isCovertAllTypeVar, command = self.toggleConvertAllType)

        yesConvertChildrenFolderRadiobutton = tk.Radiobutton(concertChildrenFolderLabelFrame, text = "转换", variable = isCovertChildrenFolderVar, value = 1)
        noConvertChildrenFolderRadiobutton = tk.Radiobutton(concertChildrenFolderLabelFrame, text = "不转换", variable = isCovertChildrenFolderVar, value = 0)

        wordCheckbutton.pack(side = tk.LEFT)
        pptCheckbutton.pack(side = tk.LEFT)
        excelCheckbutton.pack(side = tk.LEFT)
        allTypeCheckbutton.pack(side = tk.LEFT)
        yesConvertChildrenFolderRadiobutton.pack(side = tk.LEFT)
        noConvertChildrenFolderRadiobutton.pack(side = tk.LEFT)

        # startFrame
        startButton = ttk.Button(startFrame, text = '开始', command = startConvert)
        startButton.pack(side = tk.LEFT, fill = tk.X, expand = tk.YES, ipady = 1.5)

        # logListFrame
        scrollBar = tk.Scrollbar(logListFrame)
        logListText = tk.Text(logListFrame, height = 100, yscrollcommand = scrollBar.set)
        self.logListText = logListText
        scrollBar.config(command = logListText.yview)

        scrollBar.pack(side = tk.RIGHT, fill = tk.Y)
        logListText.pack(side = tk.LEFT, fill = tk.BOTH, expand = tk.YES)

    def showLog(self):
        while not logQueue.empty():
            content = logQueue.get()
            self.logListText.insert(tk.END, content + "\n")
            self.logListText.yview_moveto(1)
        self.window.after(100, self.showLog)
    
    def chooseFolderPath(self):
        return tk.filedialog.askdirectory(initialdir = os.getcwd(), title="Select file")
    
    def setFromFolderPath(self):
        folderPath = self.chooseFolderPath()
        if len(folderPath) > 0:
            self.fromFolderEntry.delete(0, tk.END)
            self.fromFolderEntry.insert(0, formatPath(folderPath))
    
    def setToFolderPath(self):
        folderPath = self.chooseFolderPath()
        if len(folderPath) > 0:
            self.toFolderEntry.delete(0, tk.END)
            self.toFolderEntry.insert(0, formatPath(folderPath))
    
    def toggleConvertAllType(self):
        flag = isCovertAllTypeVar.get() == 1
        isCovertWordVar.set(flag)
        isCovertPPTVar.set(flag)
        isCovertExcelVar.set(flag)
    
if __name__ == "__main__":
    insertLog("注意事项：")
    insertLog("1）若文件为空，PPT 报错跳过转换，Excel 对应 PDF 不能正确打开。")
    insertLog("2）若转换 PPT 时间过长，请查看是否有弹出窗口等待确认。")
    insertLog("3）在关闭 PPT 进程时，常等待20秒左右，建议先等待，若超过30秒，再前往任务管理器手动关闭该进程。")

    GUI(window)