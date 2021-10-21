# -*- coding:utf-8 -*-
import openpyxl
import settings
import os
import JsonReader


# ファイルのパス
path = JsonReader.Read_json("./settings.json")
# 学生リスト
student_list = settings.student_list
# コースのリスト
course = settings.course
# 学年のタプル
School_year = JsonReader.Read_YearIDDefault("./settings.json")
# ソートされた結果のリスト
choose_list = settings.choose_list


def setDicYear(yearName, School_year_ID):
    newDicYear = {}
    for i in range(len(yearName)):
        newDicYear[yearName[i]] = School_year_ID[i]
    return newDicYear


def load_studentlist(path):
    """Excel ファイルのロード"""
    if os.path.isfile(path) == True:
        wb = openpyxl.load_workbook(path)
        yearName = []
        School_year_ID = []
        for ws in wb.worksheets:
            year = ws.title
            yearName.append(year)
            for row in ws.rows:
                if row[0].row == 1:
                    check_header = []
                    # １行目
                    header_cells = row
                    for cell in row:
                        check_header.append(cell.value)
                    # if ['在籍番号', '氏名', 'カナ名', '電話番号', '学校名', '在籍', '面談', 'コース'] != check_header:
                    if ['在籍番号', '氏名', 'カナ名', 'コース'] != check_header:
                        return False, [], []
                else:
                    # ２行目以降
                    row_dic = {}
                    # セルの値を「key-value」で登録
                    for k, v in zip(header_cells, row):  # zip 関数(forループで複数のリストの要素を取得)
                        if not (k.value == '電話番号' or k.value == '在籍' or k.value == '面談'):
                            row_dic[k.value] = v.value
                    student_list.append(row_dic)
        student_list_keys = list(student_list[0].keys())  # Excel の表のキー取得
        for i in student_list:
            t = str(i[student_list_keys[0]])
            School_year_ID.append(t[0:2])
        School_year_ID = list(dict.fromkeys(School_year_ID))
        newDicYear = setDicYear(yearName, School_year_ID)
        JsonReader.Write_YearIDDefault("./settings.json", newDicYear)
        return True, student_list, yearName
    else:
        return False, [], []


load_studentlist("../生徒名簿.xlsx")


def setlist_course(student_list, course):
    """リストの活用"""
    student_list_keys = list(student_list[0].keys())  # Excel の表のキー取得
    for i in student_list:
        course.append(str(i[student_list_keys[3]]))
    course = list(set(course))
    sorted(course)
    return course


def setlist_keys(student_list):
    """リストの活用"""
    student_list_keys = list(student_list[0].keys())  # Excel の表のキー取得
    sorted(student_list_keys)
    return student_list_keys


def setlist_ID(student_list):
    """リストの活用"""
    School_year_ID = []
    student_list_keys = list(student_list[0].keys())  # Excel の表のキー取得
    for i in student_list:
        t = str(i[student_list_keys[0]])
        School_year_ID.append(t[0:2])
    School_year_ID = list(set(School_year_ID))
    School_year_ID.sort()
    return School_year_ID


def choose_CYname(student_list_keys, name1, name2, kana_num):  # 選択するコース,学年のリストの要素数
    """生徒の選択"""
    kana_list = [
        ["ア", "イ", "ウ", "エ", "オ", "ｱ", "ｲ", "ｳ", "ｴ", "ｵ"],
        ["カ", "キ", "ク", "ケ", "コ", "ｶ", "ｷ", "ｸ", "ｹ", "ｺ",
         "ガ", "ギ", "グ", "ゲ", "ゴ", "ｶﾞ", "ｷﾞ", "ｸﾞ", "ｹﾞ", "ｺﾞ"],
        ["サ", "シ", "ス", "セ", "ソ", "ｻ", "ｼ", "ｽ", "ｾ", "ｿ",
         "ザ", "ジ", "ズ", "ゼ", "ゾ", "ｻﾞ", "ｼﾞ", "ｽﾞ", "ｾﾞ", "ｿﾞ"],
        ["タ", "チ", "ツ", "テ", "ト", "ﾀ", "ﾁ", "ﾂ", "ﾃ", "ﾄ",
         "ダ", "ヂ", "ヅ", "デ", "ド", "ﾀﾞ", "ﾁﾞ", "ﾂﾞ", "ﾃﾞ", "ﾄﾞ"],
        ["ナ", "ニ", "ヌ", "ネ", "ノ", "ﾅ", "ﾆ", "ﾇ", "ﾈ", "ﾉ"],
        ["ハ", "ヒ", "フ", "ヘ", "ホ", "ﾊ", "ﾋ", "ﾌ", "ﾍ", "ﾎ",
         "バ", "ビ", "ブ", "ベ", "ボ", "ﾊﾞ", "ﾋﾞ", "ﾌﾞ", "ﾍﾞ", "ﾎﾞ",
         "パ", "ピ", "プ", "ペ", "ポ", "ﾊﾟ", "ﾋﾟ", "ﾌﾟ", "ﾍﾟ", "ﾎﾟ"],
        ["マ", "ミ", "ム", "メ", "モ", "ﾏ", "ﾐ", "ﾑ", "ﾒ", "ﾓ"],
        ["ヤ", "ユ", "ヨ", "ﾔ", "ﾕ", "ﾖ"],
        ["ラ", "リ", "ル", "レ", "ロ", "ﾗ", "ﾘ", "ﾙ", "ﾚ", "ﾛ"],
        ["ワ", "ヲ", "ン", "ﾜ", "ｦ", "ﾝ"]
    ]
    choose_list = []
    choose_course = name1
    choose_year = name2
    kana_list = kana_list[kana_num]
    DicYear = JsonReader.Read_YearIDDefault("./settings.json")
    for i in student_list:
        if i[student_list_keys[3]] == choose_course:
            if str(i[student_list_keys[0]]).startswith(DicYear[choose_year]):
                if (i[student_list_keys[2]])[0:1] in kana_list:
                    choose_list.append(i[student_list_keys[1]])
    choose_list = list(set(choose_list))
    sorted(choose_list)
    return choose_list
