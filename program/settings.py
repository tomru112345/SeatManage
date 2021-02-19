# -*- coding:utf-8 -*-
# デフォルトの設定を以下に記載する
# Excel ファイルのパス
Path = 'D:/Eisu_Seat_manage/生徒名簿.xlsx'

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
    #"その他" # その他を追加したときの動作を考える
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

text_set = """
        [1] 学年, [2] コースを選んで A のボタンを押してください。
        [3] 名前 の自分の名前を選んで B のボタンを押してください。
        """

LAYOUT = [
    ['1', '2', '3', '4', '5'],
    ['6', '7', '8', '9', '10'],
    ['11', '12', '13', '14', '15'],
    ['16', '17', '18', '19', 'x'],
    ['20', '21', '22', '23', '24'],
    ['25', '26', '27', '28', '29'],
    ['30', '31', '32', '33', '34']
]

License = """
        Copyright © 2001-2021 Python Software Foundation; All Rights Reserved
        
        Copyright © 2021 Kazuma Tamura
        This software is released under the MIT License
        see https://opensource.org/licenses/MIT
        """