from tkinter import *
from tkinter import filedialog
import os.path
from define import *
from mainAction import MainAction
from log import Log


orderPath = ''
invoicePath = ''
def browseFiles(num):
    filename = filedialog.askopenfilename(initialdir = "/",
                                          title = "Select a File",
                                          filetypes = (("Excel files",
                                                        "*.xlsx*"),
                                                       ("all files",
                                                        "*.*")))

    # Change label contents
    if num == ORDER_MODE:
        orderEPLabel.configure(text="찾은 파일: .../" + os.path.basename(filename))
        MainAction.setOrderPath(filename)

    elif num == INVOICE_MODE:
        invoicEPLabel.configure(text="찾은 파일: .../" + os.path.basename(filename))
        MainAction.setInvoicePath(filename)

def combineExcel():
    Log.writeLog('프로그램 실행----------------------------',__file__, 'w')
    getSheetName()
    MainAction.runMode(COMBINE_MODE)
    Log.writeLog('프로그램 종료----------------------------',__file__)

def makeExcel():
    Log.writeLog('프로그램 실행----------------------------',__file__, 'w')
    getSheetName()
    MainAction.runMode(MAKE_MODE)
    Log.writeLog('프로그램 종료----------------------------',__file__)

def addBlackList():
    pass

def setBlackList():
    pass

def getBlackList():
    pass

def getSheetName():
    # sheet이름 안적어 두면 작동 안함. 
    orderSheet = MainAction.setOrderSheet(orderSEntry.get())
    MainAction.setInvoiceSheet(invoicSEntry.get())
    


mainView = Tk()

mainView.title("엑셀 작업")
mainView.geometry(VIEW_SIZE)
mainView.config(background="white")
mainView.resizable(True, True)
versionLabel = Label(mainView, text=VERSION, fg="black")
versionLabel.place(x=440, y=160)
versionLabel.config(background="white")

# 발주서 UI부분
orderELabel = Label(mainView, text="발주서 엑셀", fg="black")
orderEPLabel = Label(mainView, width=50, text="", fg="black", anchor="w")
orderEButton = Button(mainView, text="엑셀 찾기", command=lambda: browseFiles(ORDER_MODE))
orderSLabel = Label(mainView, text="발주서 시트", fg="black")
orderSEntry = Entry(mainView)

orderELabel.grid(row=0, column=0)
orderEPLabel.grid(row=0, column=1)
orderEButton.grid(row=0, column=2)
orderSLabel.grid(row=1, column=0)
orderSEntry.grid(row=1, column=1)


# 송장 UI부분 
invoicELabel = Label(mainView, text="택배 엑셀", fg="black")
invoicEPLabel = Label(mainView, width=50, text="", fg="black", anchor="w")
invoicEButton = Button(mainView, text="엑셀 찾기", command=lambda: browseFiles(INVOICE_MODE))
invoicSLabel = Label(mainView, text="택배 시트", fg="black")
invoicSEntry = Entry(mainView)

invoicELabel.grid(row=2, column=0)
invoicEPLabel.grid(row=2, column=1)
invoicEButton.grid(row=2, column=2)
invoicSLabel.grid(row=3, column=0)
invoicSEntry.grid(row=3, column=1)


# 기능 버튼
# setBlackList = Button(mainView, text="블랙리스트 설정", command=setBlackList)
combineBtn = Button(mainView, text="송장번호 넣기", command=combineExcel)
makeBtn = Button(mainView, text="주문합치기", command=makeExcel)
# getBlackList = Button(mainView, text="블랙리스트 뽑기", command=getBlackList)
# setBlackList.place(x=190, y=120)
combineBtn.place(x=295, y=120)
makeBtn.place(x=395, y=120)
# getBlackList.place(x=390, y=120)


mainView.mainloop()