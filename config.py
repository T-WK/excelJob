from jsonManager import JsonManager
import platform
import os
from log import Log

# 운영체제 별로 설졍이 구분되어 있음.
# default: 기본설정
# custom: 사용자 설정

'''
설정 목록
downLoadPath
blackList
    {
        "name": "",
        "address": "",
        "phone": ""
    }
'''

class Config:
    def __init__(self) -> None:
        self.defaultConfig = self.getDefaultConfig()
        self.customConfig = self.getCustomConfig()

    def getConfig(self):
        if platform.system() == 'Windows':
            return JsonManager.readJson()['Windows']
        
        elif platform.system() == 'Darwin':
            return JsonManager.readJson()['Mac']

    def getDefaultConfig(self):
        return self.getConfig()['default']
    
    def getCustomConfig(self):
        return self.getConfig()['custom']


    def getDownloadPath(self):
        path = self.getCustomConfig()['downloadPath']
        if path == '':
            Log.writeLog('커스텀 경로 없음',__file__)
            path = self.getDefaultConfig()['downloadPath']

        if not os.path.isdir(path):
            Log.writeLog('경로 생성 path=%s'%path, __file__)
            os.makedirs(path)
        
        return path
        

    
    def getBlackList(self):
        if "blackList" in self.customConfig:
            return self.customConfig["blackList"]
        
        else:
            Log.writeLog('블랙리스트를 찾을 수 없음', __file__)
            return []