'''
엑셀 관련 기능 클래스
'''
import pandas as pd
import openpyxl
import os.path
import datetime as dt
from log import Log


class ExcelManager :
    
    def __init__(self) -> None:
        # self.__excelFilePath = ""
        self._errorMsg = ""
        self.__excelFileExtension = '.xlsx'
        self.__excelDownloadPath = "newExcelFile"
        self.wb = None

    # 엑셀파일 열기
    def readExcelFile(self, fileName) -> pd:
        return self.readExcelFile(fileName=fileName, sheetName=0)
    
    def readExcelFile(self, fileName, sheetName) -> pd:
        return pd.read_excel(fileName, 
                             header=1,
                             sheet_name=sheetName,
                             engine='openpyxl')
    # -

    # 엑셀파일 생성 관련 -
    def checkExtension(self, fileName) -> bool:
        if len(fileName) > 4:
            return fileName[-5:] == self.__excelFileExtension
        
    def makeDownloadPath(self, fileName) -> str:
        if not self.checkExtension(fileName=fileName):
            fileName += self.__excelFileExtension

        thisTime = dt.datetime.now().strftime("%Y%m%d:%H%M%S_")
        newFileName = thisTime + os.path.basename(fileName)
        return os.path.join(self.__excelDownloadPath, newFileName)

    def createNewExcelFile(self, dataFrame) -> None:
        newPath = self.makeDownloadPath('newExcel')
        dataFrame.to_excel(newPath, index=False)
    # -

    # 엑셀파일 수정
    def combineExcels(self, invoiceExcelName, orderExcelName):
        self.combineExcels(
            invoiceExcelName=invoiceExcelName,
            invoiceSheet=0,
            orderExcelName=orderExcelName,
            orderSheet=0    
        )

    def combineExcels(self, invoiceExcelName, invoiceSheet, orderExcelName, orderSheet):
        invoiceDf = self.readExcelFile(invoiceExcelName, invoiceSheet) # 송장 엑셀
        orderDf = self.readExcelFile(orderExcelName, orderSheet) # 발주서 엑셀

        if invoiceDf is None:
            return "invoice not found"
        
        if orderDf is None:
            return "order not found"
        
        # 수하인명 and 수하인주소 and (수하인전화번호1 or 수하인전화번호2)
        for i in range(len(invoiceDf)):
            # 조건에 맞는 행 인덱스 찾음
            name = invoiceDf.loc[i, '수하인명']
            address = invoiceDf.loc[i, '수하인주소']
            phone1 = invoiceDf.loc[i, '수하인전화번호1']
            phone2 = invoiceDf.loc[i, '수하인전화번호2']

            idx = None
            try:
                idx = orderDf.index[
                    (orderDf['수하인명'] == name)
                    &
                    (orderDf['수하인주소'] == address)
                    &
                    (orderDf['수하인전화번호1'] == phone1)
                    | 
                    (orderDf['수하인전화번호2'] == phone2)
                ].tolist()
            except Exception as e:
                self._errorMsg += str(e)
                return

            if len(idx) != 1:
                self._errorMsg += "{idx}의 행들의 수하인이 겹칩니다.\n"
                continue
            

            if '송장번호' not in orderDf.columns.values.tolist():
                orderDf['송장번호'] = pd.Series()
            # 송장 추가
            orderDf.loc[idx[0], '송장번호'] = invoiceDf.loc[i, '송장번호']

        self.createNewExcelFile(orderDf)
        
        
        

if __name__ == "__main__":
    ex = ExcelManager()
    # /Users/tw-k/Desktop/0530테스트발주서.xlsx
    ex.combineExcels(
        '/Users/tw-k/Desktop/0530테스트발주서.xlsx',
        '시트 2',
        '/Users/tw-k/Desktop/0530테스트발주서.xlsx',
        '시트 1'
    )
    Log._printError(error=ex._errorMsg)