'''
로그 작성 클래스
'''


import os.path
import datetime as dt
from define import *


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
    def writeLog(logTxt, path):
        path = os.path.basename(path)
        now = dt.datetime.now().strftime('[%H:%m:%S] - ' + path + ': ')
        txt = now + logTxt + '\n'

        fileName = Log.__getFileName()
        Log.__checkDir()

        openType = 'w'
        if Log.__checkFile(fileName=fileName):
            openType = 'a'

        with open(fileName, openType, encoding=ENCODING) as f:
            f.write(txt)
