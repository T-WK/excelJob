from excelManager import ExcelManager
from log import Log
import openManager as om
from define import *



def combineString(strArg):
    result = ''
    for val in strArg:
        result += val
    
    return result

def refine(val):
    if '"' in val:
        val = val.replace('"','')

    if "'" in val:
        val = val.replace("'",'')

    return val.strip()

def main():
    data = dict()
    ex = ExcelManager()

    Log.writeLog('json 읽기 시작.', __file__)
    with open(SETTING_FILE, 'r', encoding=ENCODING) as f:
        for line in f.readlines():
            line = line.strip()

            if '#' in line or '' == line:
                continue
            
            if len(line.split(SPLIT_STR)) == 2:
                option, val = map(refine, line.split(SPLIT_STR))
                data[option] = om.getRightPath(val)
            
            else:
                option = list(map(refine, line.split(SPLIT_STR)))[0]
                data[option] = ''
        
    Log.writeLog('실행.txt 읽기 종료.', __file__)
    if 'mode' not in data.keys():
        Log.writeLog('프로그램 모드가 존재하지 않습니다.', __file__)
        return

    data['sheet1'] = 0 if data['sheet1'] == '0' else data['sheet1']
    data['sheet2'] = 0 if data['sheet2'] == '0' else data['sheet2']

    if data['mode'] == '1':
        if 'path2' not in data.keys():
            Log.writeLog('병합할 송장 파일이 없습니다.', __file__)
            return
        
        options = set(['mode', 'path1', 'path2', 'sheet1', 'sheet2'])
        if options & set(data.keys()) != options:
            Log.writeLog('프로그램을 실행하기위한 옵션값이 부족하거나 옵션명이 올바르지 못합니다.', __file__)
            return
        ex.combineExcels(data['path1'], data['sheet1'], data['path2'], data['sheet2'])

    elif data['mode'] == '2':
        options = set(['mode', 'path1', 'sheet1'])
        if options & set(data.keys()) != options:
            Log.writeLog('프로그램을 실행하기위한 옵션값이 부족하거나 옵션명이 올바르지 못합니다.', __file__)
            return
        
        ex.getBlackLisOrder(data['path1'], data['sheet1'])


if __name__ == '__main__':
    Log.writeLog('프로그램 실행----------------------------',__file__, 'w')
    main()
    Log.writeLog('프로그램 종료----------------------------',__file__)
            
