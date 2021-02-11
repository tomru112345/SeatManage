# -*- coding:utf-8 -*-
import openpyxl
import settings

path = settings.Path # ファイルのパス

wb = openpyxl.load_workbook(path)
ws = wb.worksheets[0]

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

for row in ws.rows:
    if row[0].row == 1:
        # １行目
        header_cells = row
    else:
        # ２行目以降
        row_dic = {}
        # セルの値を「key-value」で登録
        for k, v in zip(header_cells, row): # zip 関数(forループで複数のリストの要素を取得)
            row_dic[k.value] = v.value
        student_list.append(row_dic)

student_list_keys = list(student_list[0].keys()) # Excel の表のキー取得

for i in student_list:
    t = str(i[student_list_keys[0]])
    School_year_ID.append(t[0:2])
    course.append(str(i[student_list_keys[7]]))

course = list(set(course))
School_year_ID = list(set(School_year_ID))
School_year_ID.sort()



# 参考: https://gammasoft.jp/blog/read-rows-of-excel-sheet-using-python/
"""
def choose_course(number): # 選択するコースのリストの要素数
    choose_course = course[number]
    for i in student_list:
        if i[student_list_keys[7]] == choose_course:
          choose_list.append(i)
   
def choose_year(number): # 選択する学年のリストの要素数
    choose_year = School_year_ID[number]
    for i in student_list:
        if str(i[student_list_keys[0]]).startswith(choose_year):
          choose_list.append(i) 

def choose_CY(number1, number2): # 選択するコース,学年のリストの要素数
    choose_course = course[number1]
    choose_year = School_year_ID[number2]
    for i in student_list:
        if i[student_list_keys[7]] == choose_course and str(i[student_list_keys[0]]).startswith(School_year[choose_year]):
          choose_list.append(i)   

def choose_coursename(name): # 選択するコースのリストの要素数
    choose_course = name
    for i in student_list:
        if i[student_list_keys[7]] == choose_course:
          choose_list.append(i)
   
def choose_yearname(name): # 選択する学年のリストの要素数
    choose_year = name
    for i in student_list:
        if str(i[student_list_keys[0]]).startswith(choose_year):
          choose_list.append(i) 
"""

def choose_CYname(name1, name2): # 選択するコース,学年のリストの要素数
    global choose_list
    choose_list = []
    choose_course = name1
    choose_year = name2
    for i in student_list:
        if i[student_list_keys[7]] == choose_course and str(i[student_list_keys[0]]).startswith(School_year[choose_year]):
          choose_list.append(i[student_list_keys[1]])
    return choose_list
