# -*- coding:utf-8 -*-
import json
import os
from tokenize import Ignore


def Read_json(Path):
    if os.path.isfile(Path):
        with open(Path, 'r', encoding='utf-8_sig') as f:
            data = json.load(f)
            StudentsListPath = data['FilePath']['StudentsListPath']
            return StudentsListPath


def Read_SeatDefault(Path):
    if os.path.isfile(Path):
        with open(Path, 'r', encoding='utf-8_sig') as f:
            data = json.load(f)
            height = data['SeatSettings']['height']
            length = data['SeatSettings']['length']
            Ignore_lists = data['SeatSettings']['IgnoreSeat']
            return height, length, Ignore_lists

def Read_YearIDDefault(Path):
    DicYear = {}
    if os.path.isfile(Path):
        with open(Path, 'r', encoding='utf-8_sig') as f:
            data = json.load(f)
            for i in data['Year']:
                DicYear[i] = data['Year'][i]
        return DicYear


def Write_YearIDDefault(Path, DicYear):
    if os.path.isfile(Path):
        with open(Path, 'r', encoding='utf-8_sig') as f:
            data = json.load(f)
        Year = data['Year']
        for i in DicYear.keys():
            Year[i] = DicYear[i]
        # Pythonオブジェクトをファイル書き込み
        with open(Path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)


def Write_Json(Path, value):
    if os.path.isfile(Path):
        with open(Path, 'r', encoding='utf-8_sig') as f:
            data = json.load(f)
        StudentsListPath = data['FilePath']
        StudentsListPath['StudentsListPath'] = value
        # Pythonオブジェクトをファイル書き込み
        with open(Path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)

def Write_SeatDefault(Path, height, length, Ignore_list = []):
    if os.path.isfile(Path):
        with open(Path, 'r', encoding='utf-8_sig') as f:
            data = json.load(f)
        SeatSettings = data['SeatSettings']
        SeatSettings['height'] = height
        SeatSettings['length'] = length
        tmp_ignore_list = []
        if (len(Ignore_list) > 0):
            for tmp_item in Ignore_list:
                tmp_ignore_dict = {"height": tmp_item[0], "length": tmp_item[1]}
                tmp_ignore_list.append(tmp_ignore_dict)
        SeatSettings['IgnoreSeat'] = tmp_ignore_list
        # Pythonオブジェクトをファイル書き込み
        with open(Path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)

# IgnoreList = [
#     [3,4],
#     [1,4],
#     [0,5],
# ]

# Write_SeatDefault("./settings.json", 7, 5, IgnoreList)
