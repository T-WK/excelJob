'''
로그 작성 클래스
'''


import os.path
import datetime as dt


class Log :
    __logDir = 'log'
    __f = None
    __fileName = ''

    def __init__(self, logTxt) -> None:
        self.__openLogFile()
        self.__f.write(logTxt + '\n')
        self.__closeFile()

    def __checkFile(self) -> bool:
        Log.__fileName = dt.datetime.now().strftime("%Y%m%d") + '.txt'
        Log.__fileName = os.path.join(Log.__logDir, Log.__fileName)

        if os.path.isfile(Log.__fileName):
            return True
        
        return False

    def __openLogFile(self):
        if Log.__f is not None and not Log.__f.closed():
            Log.__f.close()


        if self.__checkFile():
            Log.__f = open(Log.__fileName, 'a')
        
        else:
            Log.__f = open(Log.__fileName, 'w')
    
    def __closeFile(self):
        if Log.__f is not None and not Log.__f.closed:
            Log.__f.close()

    @classmethod
    def _writeLog(cls, logTxt):
        now = dt.datetime.now().strftime('[%H:%m:%S] - ')
        cls(now + logTxt)
