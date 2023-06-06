import platform
import os.path
    
def getWindowsPath(path):
    return path.replace('/', '\\')

def getMacPath(path):
    return path.replace('\\', '/')

def getRightPath(path):
    path = path.strip()
    if platform.system() == 'Windows':
        return getWindowsPath(path)
        
    elif platform.system() == 'Darwin':
        return getMacPath(path)

def getOtherName(path):
    for i in range(100):
        tmp = path[:] + ' (%d)'%i
        if not os.path.isfile(tmp):
            return tmp
        