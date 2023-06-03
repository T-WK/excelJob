import json
class JsonManager:
    __config = 'config/config.json'

    @staticmethod
    def readJson():
        with open(JsonManager.__config, 'r') as f:
            data = json.load(f)
        
        return data