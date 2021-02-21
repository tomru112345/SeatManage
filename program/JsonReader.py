import json
import os

def Read_json(Path):
    if os.path.isfile(Path) == True:
        with open(Path, 'r', encoding='utf-8_sig') as f:
            data = json.load(f)
            StudentsListPath = data['FilePath']['StudentsListPath']
            return StudentsListPath

def Read_SeatDefault(Path):
    if os.path.isfile(Path) == True:
        with open(Path, 'r', encoding='utf-8_sig') as f:
            data = json.load(f)
            height = data['SeatSettings']['height']
            length = data['SeatSettings']['length']
            return height, length

def Read_YearIDDefault(Path):
    School_year_ID ={}
    if os.path.isfile(Path) == True:
        with open(Path, 'r', encoding='utf-8_sig') as f:
            data = json.load(f)
            for i in data['Year']:
                School_year_ID[i] = data['Year'][i]
        return School_year_ID

def Write_Json(Path, value):
    if os.path.isfile(Path) == True:
        with open(Path, 'r', encoding='utf-8_sig') as f:
            data = json.load(f)
        StudentsListPath = data['FilePath']
        StudentsListPath['StudentsListPath'] = value
        # Pythonオブジェクトをファイル書き込み
        with open(Path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)

Read_YearIDDefault("./settings.json")


