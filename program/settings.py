# -*- coding:utf-8 -*-
# デフォルトの設定を以下に記載する
# Excel ファイルのパス
Path = "../生徒名簿.xlsx"

# 学生リスト
student_list = []

# コースのリスト
course = []

# 学年IDのリスト
School_year_ID = []

# 学年のリスト
School_year = [
    "小3",
    "小4",
    "小5",
    "小6",
    "中1",
    "中2",
    "中3"
]

School_year_tpl = {
    "小3" : "07",
    "小4" : "08",
    "小5" : "09",
    "小6" : "10",
    "中1" : "11",
    "中2" : "12",
    "中3" : "13"
}

# ソートされた結果のリスト
choose_list = []

height = 7  # 縦の席の数
length = 5  # 横の席の数

pingfile = '../image/capital_e.png'

No_Vacant_Seat = []