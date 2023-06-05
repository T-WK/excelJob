from excelManager import ExcelManager
from log import Log

SETTING_FILE = '실행.txt'
COMBINE_MODE = 1
EXTRACTION_MODE = 2


def combineString(strArg):
    result = ''
    for val in strArg:
        result += val
    
    return result

def refine(val):
    if '"' in val:
        val = combineString(list(map(str, val.split('"'))))

    if "'" in val:
        val = combineString(list(map(str, val.split("'"))))

    if '\\' in val:
        val.replace('\\', '/')
    
    return val

def main():
    data = dict()
    ex = ExcelManager()

    with open(SETTING_FILE, 'r') as f:
        for line in f.readlines():
            line = line.strip()

            if '#' in line or '' == line:
                continue
            
            if len(line.split(':')) == 2:
                option, val = map(refine, line.split(':'))
                data[option] = val
            
            else:
                option = list(map(refine, line.split(':')))[0]
                data[option] = ''

    if 'mode' not in data.keys():
        Log.writeLog('프로그램 모드가 존재하지 않습니다.', __file__)
        return

    if int(data['mode']) == 1:
        if 'path2' not in data.keys():
            Log.writeLog('병합할 송장 파일이 없습니다.', __file__)
            return
        
        options = set(['mode', 'path1', 'path2', 'sheet1', 'sheet2'])
        if options & set(data.keys()) != options:
            Log.writeLog('프로그램을 실행하기위한 옵션값이 부족하거나 옵션명이 올바르지 못합니다.', __file__)
            return
        ex.combineExcels(data['path1'], data['sheet1'], data['path2'], data['sheet2'])

    elif int(data['mode']) == 2:
        options = set(['mode', 'path1', 'sheet1'])
        if options & set(data.keys()) != options:
            Log.writeLog('프로그램을 실행하기위한 옵션값이 부족하거나 옵션명이 올바르지 못합니다.', __file__)
            return
        
        ex.getBlackLisOrder(data['path1'], data['sheet1'])


if __name__ == '__main__':
    main()
            