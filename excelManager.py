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
                             engine='openpyxl')
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
            name = invoiceDf.loc[i, '수하인명']
            address = invoiceDf.loc[i, '수하인주소']
            phone1 = invoiceDf.loc[i, '수하인 전화번호 1']
            phone2 = invoiceDf.loc[i, '수하인 전화번호 2']

            nameBSr = orderDf['수하인명'] == invoiceDf.loc[i, '수하인명']
            addressBSr = orderDf['수하인주소'] == invoiceDf.loc[i, '수하인주소']
            phone1BSr = orderDf['수하인 전화번호 1'] == invoiceDf.loc[i, '수하인 전화번호 1']
            phone2BSr = orderDf['수하인 전화번호 2'] == invoiceDf.loc[i, '수하인 전화번호 2']

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
                        j, invoiceDf.loc[i, '수하인명'], invoiceDf.loc[i, '수하인주소'],
                        invoiceDf.loc[i, '수하인 전화번호 1'], invoiceDf.loc[i, '수하인 전화번호 2']), __file__)

                continue
            

            if '송장번호' not in orderDf.columns.values.tolist():
                orderDf['송장번호'] = pd.Series()
            # 송장 추가
            orderDf.loc[idx[0], '송장번호'] = invoiceDf.loc[i, '송장번호']

        self.createNewExcelFile(orderDf, 'combindedExcel')
    
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