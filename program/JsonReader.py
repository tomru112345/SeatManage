import json
import os

def Read_json(Path):
    if os.path.isfile(Path) == True:
        with open(Path, 'r', encoding='utf-8_sig') as f:
            data = json.load(f)
            StudentsListPath = data['FilePath']['StudentsListPath']
            return StudentsListPath

def Write_Json(Path, value):
    if os.path.isfile(Path) == True:
        with open(Path, 'r', encoding='utf-8_sig') as f:
            data = json.load(f)
        StudentsListPath = data['FilePath']
        StudentsListPath['StudentsListPath'] = value
        # Pythonオブジェクトをファイル書き込み
        with open(Path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)


