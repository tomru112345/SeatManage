# -*- coding:utf-8 -*-
import datetime
import openpyxl
import os
import settings

Seat = [] # リスト型で保存

No_Vacant_Seat = settings.No_Vacant_Seat

def time():
    dt_now = datetime.datetime.now()
    day_today = dt_now.strftime('%Y%m%d')
    time_now = dt_now.strftime('%H:%M:%S')
    return day_today, time_now

def New_day_book():
    day_today, time_now = time()
    if os.path.isfile(f'../log/Seat_{day_today}.xlsx') == False:
        book = openpyxl.Workbook(f'../log/Seat_{day_today}.xlsx')
        book.save(f'../log/Seat_{day_today}.xlsx')
        book.close()
        book = openpyxl.load_workbook(f'../log/Seat_{day_today}.xlsx')
        active_sheet = book.worksheets[0]
        active_sheet.append(["名前", "席番号", "日付", "入室時間", "退室時間"])
        book.save(f'../log/Seat_{day_today}.xlsx')
        book.close()

def Open_book():
    day_today, time_now = time()
    if os.path.isfile(f'../log/Seat_{day_today}.xlsx') == True:
        book = openpyxl.load_workbook(filename=f'../log/Seat_{day_today}.xlsx')
        active_sheet = book.active
        size = len(Seat)
        t = 0
        for row in active_sheet.rows:
            if row[0].row == 1:
                # １行目
                header_cells = row
            else:
                # ２行目以降
                row_dic = {}
                # セルの値を「key-value」で登録
                for k, v in zip(header_cells, row): # zip 関数(forループで複数のリストの要素を取得)
                    row_dic[k.value] = v.value
                if t < size:
                    Seat[t] =row_dic
                    t = t + 1
                else:
                    Seat.append(row_dic)

        for i in Seat:
            if i["退室時間"] == None:
                No_Vacant_Seat.append(i["席番号"])
        book.save(f'../log/Seat_{day_today}.xlsx')
        book.close()
    return No_Vacant_Seat

def append_time_in(name, number):
    day_today, time_now = time()
    book = openpyxl.load_workbook(filename=f'../log/Seat_{day_today}.xlsx')
    active_sheet = book.worksheets[0]
    active_sheet.append([name, number, day_today, time_now, ""])
    book.save(f'../log/Seat_{day_today}.xlsx')
    book.close()

def leave_seat_time(number):
    day_today, time_now = time()
    book = openpyxl.load_workbook(filename=f'../log/Seat_{day_today}.xlsx')
    active_sheet = book.active
    size = len(Seat)
    t = 0
    for row in active_sheet.rows:
        if row[0].row == 1:
            # １行目
            header_cells = row
        else:
            # ２行目以降
            row_dic = {}
            # セルの値を「key-value」で登録
            for k, v in zip(header_cells, row): # zip 関数(forループで複数のリストの要素を取得)
                row_dic[k.value] = v.value
            if t < size:
                Seat[t] =row_dic
                t = t + 1
            else:
                Seat.append(row_dic)

    for i in Seat:
        if i["席番号"] == number:
            if i["退室時間"] == None:
                active_sheet.cell(column=5, row= Seat.index(i) + 2, value= time_now)
    
    book.save(f'../log/Seat_{day_today}.xlsx')
    book.close()



