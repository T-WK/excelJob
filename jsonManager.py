import json
class JsonManager:
    __config = 'config/config.json'

    @staticmethod
    def readJson():
        with open(JsonManager.__config, 'rt', encoding='utf8') as f:
            data = json.load(f)
        
        return data