import json
from define import *
class JsonManager:
    __config = 'config/config.json'

    @staticmethod
    def readJson():
        with open(JsonManager.__config, 'rt', encoding=ENCODING) as f:
            data = json.load(f)
        
        return data