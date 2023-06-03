from jsonManager import JsonManager
import platform

# 운영체제 별로 설졍이 구분되어 있음.
# default: 기본설정
# custom: 사용자 설정

class Config:
    def getConfig(self):
        if platform.system() == 'Windows':
            return JsonManager.readJson()['Windows']
        
        elif platform.system() == 'Darwin':
            return JsonManager.readJson()['Mac']

    @staticmethod
    def getDefaultConfig(self):
        
        return self.getConfig()['default']
    
    @staticmethod
    def getCustomConfig(self):
        return self.getConfig()['custom']
