'''
로그 작성 클래스
'''


import os.path
import datetime as dt


class Log :
    __logDir = 'log'
    __f = None
    __fileName = ''

    def _checkFile(self) -> bool:
        self.__fileName = dt.datetime.now().strftime("%Y%m%d") + '.txt'
        self.__fileName = os.path.join(self.__logDir, self.__fileName)

        if os.path.isfile(self.__fileName):
            return True
        
        return False

    def _openLogFile(self):
        if not self.__f.closed():
            self.__f.close()


        if self.checkFile():
            self.__f = open(self.__fileName, 'a')
        
        else:
            self.__f = open(self.__fileName, 'w')

    def _closeFile(self):
        if not self.__f.closed():
            self.__f.close()

    def _writeLog(self, logTxt):
        self.openLogFile()
        self.__f.write(logTxt + '\n')
        self.closeFile()

        