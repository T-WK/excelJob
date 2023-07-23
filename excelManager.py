'''
엑셀 관련 기능 클래스
'''
import pandas as pd
import os.path
import datetime as dt
from log import Log
from config import Config
from define import *


class ExcelManager :
    
    def __init__(self) -> None:
        self.conf = Config()
        self.__excelDownloadPath = self.conf.getDownloadPath()
        self.__blackList = self.conf.getBlackList()
        self.__excelFileExtension = '.xlsx'

    # 엑셀파일 열기
    def readExcelFile(self, fileName) -> pd:
        return self.readExcelFile(fileName=fileName, sheetName=0)
    
    def readExcelFile(self, fileName, sheetName) -> pd:
        if not self.checkExtension(fileName):
            Log.writeLog('읽으려하는 파일의 확장자가 잘못되었습니다. fileName = %s'%fileName, __file__)
            return None
        
        return pd.read_excel(fileName, 
                             header=0,
                             sheet_name=sheetName,
                             engine='openpyxl').fillna('-')
        
    # -

    # 엑셀파일 생성 관련 -
    def checkExtension(self, fileName) -> bool:
        if len(fileName) > 4:
            return fileName[-5:] == self.__excelFileExtension
        
        return False
        
    def makeDownloadPath(self, fileName) -> str:
        if not self.checkExtension(fileName=fileName):
            fileName += self.__excelFileExtension

        thisTime = dt.datetime.now().strftime("%Y%m%d_%H%M%S_")
        newFileName = thisTime + os.path.basename(fileName)
        return os.path.join(self.__excelDownloadPath, newFileName)

    def createNewExcelFile(self, dataFrame, newName) -> None:
        newPath = self.makeDownloadPath(newName)
        dataFrame.to_excel(newPath, index=False)
    # -

    # 엑셀파일 작업
    def combineExcels(self, orderExcelName='', orderSheet=0, invoiceExcelName='', invoiceSheet=0):
        Log.writeLog('combineExcels 시작', __file__)
        invoiceDf = self.readExcelFile(invoiceExcelName, invoiceSheet) # 송장 엑셀
        orderDf = self.readExcelFile(orderExcelName, orderSheet) # 발주서 엑셀

        if invoiceDf is None or orderDf is None :
            Log.writeLog('송장DF:%d, 발주서DF:%d'%(invoiceDf==None, orderDf==None), __file__)
            return
        
        # 수하인명 and 수하인주소 and (수하인전화번호1 or 수하인전화번호2)
        for i in range(len(invoiceDf)):
            # 조건에 맞는 행 인덱스 찾음
            name = invoiceDf.loc[i, COLUMN_VALUES['NAME']]
            address = invoiceDf.loc[i, COLUMN_VALUES['ADDRESS']]
            phone1 = invoiceDf.loc[i, COLUMN_VALUES['PHONE1']]
            phone2 = invoiceDf.loc[i, COLUMN_VALUES['PHONE2']]

            nameBSr = orderDf[COLUMN_VALUES['NAME']] == invoiceDf.loc[i, COLUMN_VALUES['NAME']]
            addressBSr = orderDf[COLUMN_VALUES['ADDRESS']] == invoiceDf.loc[i, COLUMN_VALUES['ADDRESS']]
            phone1BSr = orderDf[COLUMN_VALUES['PHONE1']] == invoiceDf.loc[i, COLUMN_VALUES['PHONE1']]
            phone2BSr = orderDf[COLUMN_VALUES['PHONE2']] == invoiceDf.loc[i, COLUMN_VALUES['PHONE2']]

            idx = None
            errorMsg = ''
            idx = orderDf[nameBSr & addressBSr & phone1BSr & phone2BSr].index

            if len(idx) == 0:
                # TODO: Tlqkf 엑셀상엔 존재한느데 없다함
                Log.writeLog('%d"%s,%s,%s,%s"를 찾지 못했습니다.'%(i,name,address,phone1,phone2), __file__)
                continue

            elif len(idx) != 1:
                Log.writeLog("송장의 %d번째 행과 발주서의 [%s] 행들의 수하인이 겹칩니다.\n"%(i, ', '.join(map(str, idx))), __file__)
                Log.writeLog('송장 %d번 행 "%s,%s,%s,%s"'%(i, name,address,phone1,phone2), __file__)
                for j in idx:
                    Log.writeLog('발주서 %d번 행 "%s,%s,%s,%s"'%(
                        j, invoiceDf.loc[i, COLUMN_VALUES['수하인명']], invoiceDf.loc[i, COLUMN_VALUES['ADDRESS']],
                        invoiceDf.loc[i, COLUMN_VALUES['PHONE1']], invoiceDf.loc[i, COLUMN_VALUES['PHONE2']]), __file__)

                continue
            

            if COLUMN_VALUES['INVOICE_NUM'] not in orderDf.columns.values.tolist():
                orderDf[COLUMN_VALUES['INVOICE_NUM']] = pd.Series()
            # 송장 추가
            orderDf.loc[idx[0], COLUMN_VALUES['INVOICE_NUM']] = invoiceDf.loc[i, COLUMN_VALUES['INVOICE_NUM']]

        self.createNewExcelFile(orderDf, 'combindedExcel')

    def makeExcel(self, orderExcelName='', orderSheet=0):
        Log.writeLog('makeExcels 시작', __file__)
        orderDf = self.readExcelFile(orderExcelName, orderSheet) # 발주서 엑셀

        newDF = pd.DataFrame(columns=self.makeColumnList())
        doneIndex = []

        if orderDf is None :
            Log.writeLog('발주서DF:%d'%(orderDf==None), __file__)
            return
        
        # 수하인명 and 수하인주소 and (수하인전화번호1 or 수하인전화번호2)
        for i in range(len(orderDf)):
            # 조건에 맞는 행 인덱스 찾음
            name = orderDf.loc[i, COLUMN_VALUES['NAME']]
            address = orderDf.loc[i, COLUMN_VALUES['ADDRESS']]
            phone1 = orderDf.loc[i, COLUMN_VALUES['PHONE1']]
            phone2 = orderDf.loc[i, COLUMN_VALUES['PHONE2']]

            nameBSr = orderDf[COLUMN_VALUES['NAME']] == orderDf.loc[i, COLUMN_VALUES['NAME']]
            addressBSr = orderDf[COLUMN_VALUES['ADDRESS']] == orderDf.loc[i, COLUMN_VALUES['ADDRESS']]
            phone1BSr = orderDf[COLUMN_VALUES['PHONE1']] == orderDf.loc[i, COLUMN_VALUES['PHONE1']]
            phone2BSr = orderDf[COLUMN_VALUES['PHONE2']] == orderDf.loc[i, COLUMN_VALUES['PHONE2']]

            idx = None
            errorMsg = ''
            idx = orderDf[nameBSr & addressBSr & phone1BSr & phone2BSr].index.tolist()

            if len(idx) == 0:
                Log.writeLog('%d"%s,%s,%s,%s"를 찾지 못했습니다.'%(i,name,address,phone1,phone2), __file__)
                continue

            elif len(idx) > 0:
                if (idx[0] in doneIndex): continue

                newRowDict = self.makeRowToDict()
                orderNum = ''
                for j in idx:
                    # 값 저장해둠
                    newRowDict = self.setValueTODict(newRowDict, orderDf.loc[j])


                    # 주문번호 최소값으로 찾음
                    # if orderNum == '':
                    #     orderNum = orderDf.loc[j, COLUMN_VALUES['INVOICE_NUM']]
                    # else:
                    #     orderNum = min(int(orderDf.loc[j, COLUMN_VALUES['INVOICE_NUM']]), int(orderNum))
            
                doneIndex += idx

                newDF = pd.concat([newDF, pd.DataFrame(newRowDict)]).reset_index(drop=True)

        self.createNewExcelFile(newDF, 'combindedExcel')
    
    def makeRowToDict(self):
        newDict = dict()
        for key, value in COLUMN_VALUES.items():
            newDict[value] = ['-']

        return newDict

    def setValueTODict(self, dic, df):
        if dic[COLUMN_VALUES['NAME']] == ['-']:
            dic[COLUMN_VALUES['NAME']] = [str(df[COLUMN_VALUES['NAME']])]
            dic[COLUMN_VALUES['BLANK']] = [str(df[COLUMN_VALUES['BLANK']])]
            dic[COLUMN_VALUES['ADDRESS']] = [str(df[COLUMN_VALUES['ADDRESS']])]
            dic[COLUMN_VALUES['PHONE1']] = [str(df[COLUMN_VALUES['PHONE1']])]
            dic[COLUMN_VALUES['PHONE2']] = [str(df[COLUMN_VALUES['PHONE2']])]
            dic[COLUMN_VALUES['COUNT']] = ['1']
            dic[COLUMN_VALUES['ITEM']] = [str(df[COLUMN_VALUES['ITEM']])]
            dic[COLUMN_VALUES['OPTION']] = [str(df[COLUMN_VALUES['OPTION']])]
            dic[COLUMN_VALUES['OPTION1']] = [str(df[COLUMN_VALUES['OPTION1']])]
            dic[COLUMN_VALUES['OPTION2']] = [str(df[COLUMN_VALUES['OPTION2']])]
            dic[COLUMN_VALUES['DELIVERY_MSG']] = [str(df[COLUMN_VALUES['DELIVERY_MSG']])]
            dic[COLUMN_VALUES['MALL']] = [str(df[COLUMN_VALUES['MALL']])]
            dic[COLUMN_VALUES['ORDER_NUM']] = [str(df[COLUMN_VALUES['ORDER_NUM']])]
        
        else:
            dic[COLUMN_VALUES['ITEM']][0] += '\n' + str(df[COLUMN_VALUES['ITEM']])
            dic[COLUMN_VALUES['OPTION']] = [str(df[COLUMN_VALUES['OPTION']])]
            dic[COLUMN_VALUES['OPTION1']][0] += '\n' + str(df[COLUMN_VALUES['OPTION1']])
            dic[COLUMN_VALUES['OPTION2']][0] += '\n' + str(df[COLUMN_VALUES['OPTION2']])
            
            # dic[COLUMN_VALUES['ORDER_NUM']] = [str(min(int(dic[COLUMN_VALUES['ORDER_NUM']][0]), int(df[COLUMN_VALUES['ORDER_NUM']])))]
            if str(dic[COLUMN_VALUES['ORDER_NUM']][0]).isdigit() and str(df[COLUMN_VALUES['ORDER_NUM']]).isdigit():
                dic[COLUMN_VALUES['ORDER_NUM']] = [str(min(int(dic[COLUMN_VALUES['ORDER_NUM']][0]), int(df[COLUMN_VALUES['ORDER_NUM']])))]
            
            elif not str(dic[COLUMN_VALUES['ORDER_NUM']][0]).isdigit() and str(df[COLUMN_VALUES['ORDER_NUM']]).isdigit():
                dic[COLUMN_VALUES['ORDER_NUM']] = [df[COLUMN_VALUES['ORDER_NUM']]]
            
            elif str(dic[COLUMN_VALUES['ORDER_NUM']][0]).isdigit() and not str(df[COLUMN_VALUES['ORDER_NUM']]).isdigit():
                dic[COLUMN_VALUES['ORDER_NUM']] = [dic[COLUMN_VALUES['ORDER_NUM']][0]]

        return dic

    def makeColumnList(self):
        newList = []
        for key, value in COLUMN_VALUES.items():
            newList.append(value)

        return newList

    def getBlackLisOrder(self, orderExcel, sheetName):
        Log.writeLog('getBlackLisOrder 시작', __file__)
        if self.__blackList == []:
            Log.writeLog('블랙리스트 목록이 비어있음', __file__)
            return

        orderDf = self.readExcelFile(orderExcel, sheetName)
        for info in self.__blackList:
            name = info['name']
            address = info['address']
            phone = info['phone']

            indexArr = []
            idx = orderDf.index[
                (orderDf['수하인명'] == name)
                &
                (orderDf['수하인주소'] == address)
                &
                (orderDf['수하인 전화번호 1'] == phone)
                | 
                (orderDf['수하인 전화번호 2'] == phone)
            ].tolist()
            
            indexArr += idx
        
        blackListDf = orderDf.iloc[indexArr, :]
        self.createNewExcelFile(blackListDf, 'blackList')
