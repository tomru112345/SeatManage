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
# 学年IDのリスト
School_year_ID = settings.School_year_ID
# 学年のリスト
School_year = settings.School_year_tpl
# ソートされた結果のリスト
choose_list = settings.choose_list


def load_studentlist(path):
    """Excel ファイルのロード"""
    if os.path.isfile(path) == True:
        wb = openpyxl.load_workbook(path)
        for ws in wb.worksheets:
            for row in ws.rows:
                if row[0].row == 1:
                    check_header = []
                    # １行目
                    header_cells = row
                    for cell in row:
                        check_header.append(cell.value)
                    if ['在籍番号', '氏名', 'カナ名', '電話番号', '学校名', '在籍', '面談', 'コース'] != check_header:
                        return False, []
                else:
                    # ２行目以降
                    row_dic = {}
                    # セルの値を「key-value」で登録
                    for k, v in zip(header_cells, row): # zip 関数(forループで複数のリストの要素を取得)
                        row_dic[k.value] = v.value
                    student_list.append(row_dic)
        return True, student_list

def setlist_course(student_list, course):
    """リストの活用"""
    student_list_keys = list(student_list[0].keys()) # Excel の表のキー取得
    for i in student_list:
        course.append(str(i[student_list_keys[7]]))
    course = list(set(course))
    sorted(course)
    return course

def setlist_keys(student_list):
    """リストの活用"""
    student_list_keys = list(student_list[0].keys()) # Excel の表のキー取得
    sorted(student_list_keys)
    return student_list_keys

def setlist_ID(student_list, School_year_ID):
    """リストの活用"""
    student_list_keys = list(student_list[0].keys()) # Excel の表のキー取得
    for i in student_list:
            t = str(i[student_list_keys[0]])
            School_year_ID.append(t[0:2])
    School_year_ID = list(set(School_year_ID))
    School_year_ID.sort()
    return School_year_ID

def choose_CYname(student_list_keys, name1, name2): # 選択するコース,学年のリストの要素数
    """生徒の選択"""
    choose_list = []
    choose_course = name1
    choose_year = name2
    for i in student_list:
        if i[student_list_keys[7]] == choose_course and str(i[student_list_keys[0]]).startswith(School_year[choose_year]):
          choose_list.append(i[student_list_keys[1]])
    choose_list = list(set(choose_list))
    sorted(choose_list)
    return choose_list

