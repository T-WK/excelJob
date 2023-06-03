'''
로그 작성 클래스
'''


import os.path
import datetime as dt


class Log :
    __logDir = 'log'

    def __getFileName() -> str:
        fileName = dt.datetime.now().strftime("%Y%m%d") + '.txt'
        fileName = os.path.join(Log.__logDir, fileName)
        return fileName
    
    def __checkFile(fileName) -> bool:
        if os.path.isfile(fileName):
            return True
        
        return False
    
    def __checkDir():
        if not os.path.isdir(Log.__logDir):
            os.mkdir(Log.__logDir)

    @staticmethod
    def writeLog(logTxt):
        now = dt.datetime.now().strftime('[%H:%m:%S] - ')
        txt = now + logTxt

        fileName = Log.__getFileName()
        Log.__checkDir()

        openType = 'w'
        if Log.__checkFile(fileName=fileName):
            openType = 'a'

        with open(fileName, openType) as f:
            f.write(txt)
