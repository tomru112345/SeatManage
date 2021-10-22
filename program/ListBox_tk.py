# -*- coding:utf-8 -*-
from tkinter import *
import os
import tkinter.filedialog
from tkinter import font
from tkinter import ttk
import Sort_Students
import settings
import SeatNumber
import JsonReader
import importlib
from tkinter import messagebox
import sys

height, length = JsonReader.Read_SeatDefault("./settings.json")

pingfile = settings.pingfile

Select_Student = ""
Select_Number = 0
select_course = ""
select_year = ""
time = ""

path = JsonReader.Read_json("./settings.json")  # ファイルのパス

# 学生リスト
student_list = settings.student_list
# コースのリスト
course = settings.course
# 学年IDのリスト
School_year_ID = settings.School_year_ID
# 学年のタプル
School_year = JsonReader.Read_YearIDDefault("./settings.json")
# ソートされた結果のリスト
choose_list = settings.choose_list
student_list_keys = []

font_name = settings.font_name

# 2次元配列のとおりに、gridでレイアウトを作成する

LAYOUT = settings.LAYOUT

buttons = []


class Seat(ttk.Frame):  # リストボックスのクラス

    def __init__(self, master=None):
        """初期化"""
        super().__init__(master)
        self.create_style()
        self.create_widgets()

    def reload_modules(self):
        """リロード"""
        importlib.reload(settings)
        importlib.reload(Sort_Students)
        importlib.reload(SeatNumber)

    def create_style(self):
        """ボタン、ラベルのスタイルを変更."""
        style = ttk.Style()
        style.theme_use('alt')
        # ボタンのスタイルを上書き
        style.configure('MyWidget.TButton', font=(
            font_name, 20), background='#32CD32')

        style2 = ttk.Style()
        style2.theme_use('alt')
        # ボタンのスタイルを上書き
        style2.configure('office.TButton', font=(
            font_name, 20), background='#D3D3D3')

        style3 = ttk.Style()
        style3.theme_use('alt')
        # ボタンのスタイルを上書き
        style3.configure('MyWidget2.TButton', font=(
            font_name, 20), background='#DC143C')

        style4 = ttk.Style()
        style4.theme_use('alt')
        # ボタンのスタイルを上書き
        style4.configure('office2.TButton', font=(
            font_name, 10), background='#D3D3D3')

    def create_widgets(self):
        """席ボタンウィジェットの作成."""
        font0 = font.Font(family=font_name, size=20, weight='bold')
        self.label0 = ttk.Label(self, text="""
        自習室の希望する席を選んでください。
        * 緑 : 席が空いてます
        * 赤 : 席を使っています
        """, font=font0, anchor='e', justify='left')
        self.label0.grid(column=0, row=0, columnspan=5)
        # レイアウトの作成
        for y, row in enumerate(LAYOUT, 1):
            for x, char in enumerate(row):
                if char != "x":
                    No_Vacant_Seat = SeatNumber.Open_book()
                    if int(char) in No_Vacant_Seat:
                        index = int(char) - 1
                        buttons.append(ttk.Button(
                            self, text=char, style='MyWidget2.TButton'))
                        buttons[index].grid(
                            column=x, row=y, sticky=(N, S, E, W))
                        buttons[index].bind(
                            '<Button-1>', func=self.click_option)
                    else:
                        index = int(char) - 1
                        buttons.append(ttk.Button(
                            self, text=char, style='MyWidget.TButton'))
                        buttons[index].grid(
                            column=x, row=y, sticky=(N, S, E, W))
                        buttons[index].bind(
                            '<Button-1>', func=self.click_option)
        self.grid(column=0, row=0, sticky=(N, S, E, W))

        # 横の引き伸ばし設定
        for i in range(length):
            self.columnconfigure(i, weight=1)

        # 縦の引き伸ばし設定。0番目の結果表示欄だけ、元の大きさのまま
        self.rowconfigure(0, weight=0)
        for i in range(height):
            self.rowconfigure(i + 1, weight=1)

        # ウィンドウ自体の引き伸ばし設定
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)

        # menubarの大元（コンテナ）の作成と設置
        menubar = Menu(self)
        file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="設定", menu=file_menu)

        # 生徒名簿の設定
        file_menu.add_command(
            label="生徒名簿", command=self.onOpenSettingStudentfile, accelerator="Ctrl+O")
        self.master.config(menu=menubar)
        self.bind_all("<Control-o>", self.onOpenSettingStudentfile)

        file_menu.add_command(
            label="終了", command=self.ExitApp, accelerator="Ctrl+F")
        self.master.config(menu=menubar)
        self.bind_all("<Control-f>", self.ExitApp)

        # ライセンス表示
        License_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="表示", menu=License_menu)
        License_menu.add_command(
            label="ライセンス", command=self.onOpenLicense, accelerator="Ctrl+L")
        self.master.config(menu=menubar)
        self.bind_all("<Control-l>", self.onOpenLicense)

        # 学年ID の設定
        License_menu.add_command(
            label="学年 ID", command=self.onOpenSettingID, accelerator="Ctrl+I")
        self.master.config(menu=menubar)
        self.bind_all("<Control-i>", self.onOpenSettingID)

    def ExitApp(self, event=None):
        check_Fin = messagebox.askyesno(
            title=f"アプリケーション終了",
            message=f"アプリケーションを終了しますか？")

        if check_Fin:
            self.reload_modules()
            sys.exit()

    def onOpenSettingID(self, event=None):
        """学年 ID"""
        self.reload_modules()
        self.dialog = Toplevel(self)
        self.dialog.title(f"学年 ID")
        window_width = 350
        window_height = 700
        x = int(int(self.dialog.winfo_screenwidth()/2) - int(window_width/2))
        y = int(int(self.dialog.winfo_screenheight()/2) -
                int(window_height/2))
        self.dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.dialog.resizable(width=False, height=False)
        self.dialog.grab_set()

        Bool_value, list_value, yearName = Sort_Students.load_studentlist(
            JsonReader.Read_json("./settings.json"))
        if Bool_value == True:
            DicYear = JsonReader.Read_YearIDDefault("./settings.json")
            font1 = font.Font(size=10, weight='bold')
            self.labellist = []
            t = 0
            self.labellist.append(Label(self.dialog, text=f"""
                現在の設定は以下になります。""", font=font1, anchor='e', justify='left'))
            for i in DicYear.keys():
                t = t + 1
                self.labellist.append(Label(self.dialog, text=f"""
                * {i} : {DicYear[i]}
                """, font=font1, anchor='e', justify='left'))
            for i in range(len(self.labellist)):
                self.labellist[i].grid(column=0, row=i)

    def onOpenLicense(self, event=None):
        """ライセンス"""
        messagebox.showinfo(
            title="ライセンス",
            message="ライセンス",
            detail=settings.License)

    def onOpenSettingStudentfile(self, event=None):
        """生徒名簿の設定"""
        self.reload_modules()
        self.dialog = Toplevel(self)
        self.dialog.title(f"生徒名簿の設定")
        window_width = 480
        window_height = 150
        x = int(int(self.dialog.winfo_screenwidth()/2) - int(window_width/2))
        y = int(int(self.dialog.winfo_screenheight()/2) -
                int(window_height/2))
        self.dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.dialog.resizable(width=False, height=False)
        self.dialog.grab_set()

        font1 = font.Font(family=font_name, size=10, weight='bold')
        self.label1 = Label(self.dialog, text=f"""
        生徒名簿の Excel ファイルを指定してください。
        """, font=font1, anchor='e', justify='left')
        self.label1.grid(column=0, row=1, columnspan=2)
        button_file = ttk.Button(
            self.dialog, text="開く", style="office2.TButton")
        button_file.bind('<Button-1>', func=self.file_dialog)
        button_file.grid(column=2, row=1)

        self.file_name = StringVar()
        self.file_name.set(JsonReader.Read_json("./settings.json"))
        self.label2 = Label(
            self.dialog, textvariable=self.file_name, font=(font_name, 10))
        self.label2.grid(column=0, row=2, columnspan=3)

        self.label2 = Label(self.dialog, textvariable="",
                            font=(font_name, 10))
        self.label2.grid(column=0, row=3, columnspan=3)

        button_fin = ttk.Button(self.dialog, text="決定",
                                style="office2.TButton")
        button_fin.bind('<Button-1>', func=self.select_filename)
        button_fin.grid(column=0, row=5, columnspan=3)

    def file_dialog(self, event):
        """ファイルの選択オプション"""
        fTyp = [("Excel", "xlsx")]
        iDir = os.path.abspath(os.path.dirname(__file__))
        file_name = tkinter.filedialog.askopenfilename(
            filetypes=fTyp, initialdir=iDir)
        if len(file_name) == 0:
            self.file_name.set(JsonReader.Read_json("./settings.json"))
        else:
            self.file_name.set(file_name)

    def select_filename(self, event):
        filename = self.file_name.get()
        JsonReader.Write_Json("./settings.json", filename)
        self.dialog.destroy()

    def click_option(self, event):
        global Select_Number
        Select_Number = int(event.widget.cget("text"))
        if (buttons[Select_Number - 1])['style'] == 'MyWidget.TButton':
            self.OpenListbox()
        elif (buttons[Select_Number - 1])['style'] == 'MyWidget2.TButton':
            self.OpenDialog()

    def OpenListbox(self):  # 席を取るときのダイアログ
        """リストウィジェットの作成"""
        self.reload_modules()
        Bool_value, list_value, yearName = Sort_Students.load_studentlist(
            JsonReader.Read_json("./settings.json"))

        if Bool_value == True:
            School_year_ID = Sort_Students.setlist_ID(list_value)

            course = Sort_Students.setlist_course(list_value, settings.course)

            self.reload_modules()
            self.dialog = Toplevel(self)

            self.dialog.title("生徒リスト")
            window_width = 960
            window_height = 960
            x = int(int(self.dialog.winfo_screenwidth()/2) - int(window_width/2))
            y = int(int(self.dialog.winfo_screenheight()/2) -
                    int(window_height/2))
            self.dialog.geometry(f"{window_width}x{window_height}+{x}+{0}")
            self.dialog.grab_set()

            # 横の引き伸ばし設定
            for i in range(3):
                self.dialog.columnconfigure(i, weight=1)

            # 縦の引き伸ばし設定。0番目の結果表示欄だけ、元の大きさのまま
            self.rowconfigure(0, weight=0)
            for i in range(7):
                self.dialog.rowconfigure(i, weight=1)

            font1 = font.Font(family=font_name, size=15, weight='bold')
            self.label1 = Label(self.dialog, text=settings.text_set,
                                font=font1, anchor='e', justify='left')
            self.label1.grid(column=0, row=0, columnspan=3)

            font2 = font.Font(family=font_name, size=15, weight='bold')
            self.label2 = Label(self.dialog, text="学年", font=font2)
            self.label2.grid(column=0, row=1, sticky=W + E + N + S)

            font3 = font.Font(family=font_name, size=15, weight='bold')
            self.label3 = Label(self.dialog, text="コース", font=font3)
            self.label3.grid(column=1, row=1, sticky=W + E + N + S)

            self.label3 = Label(self.dialog, text="カナ行", font=font3)
            self.label3.grid(column=2, row=1, sticky=W + E + N + S)

            self.year = yearName  # 学年リスト
            yearname = StringVar(value=self.year)  # 文字列を保持させる
            selectyearname = StringVar()  # 文字列を保持させる

            self.listyearbox = Listbox(self.dialog, listvariable=yearname, height=7, exportselection=0, font=(
                font_name, 15, "bold"))  # リストボックスに追加
            self.listyearbox.grid(column=0, row=2, sticky=W + E + N + S)

            self.course = course  # コースリスト
            coursename = StringVar(value=self.course)  # 文字列を保持させる

            self.listcoursebox = Listbox(self.dialog, listvariable=coursename, height=7, exportselection=0, font=(
                font_name, 15, "bold"))  # リストボックスに追加
            self.listcoursebox.grid(column=1, row=2, sticky=W + E + N + S)

            self.kana = ["ア行", "カ行", "サ行", "タ行", "ナ行",
                         "ハ行", "マ行", "ヤ行", "ラ行", "ワ行"]  # コースリスト
            kana = StringVar(value=self.kana)  # 文字列を保持させる

            self.listkanabox = Listbox(self.dialog, listvariable=kana, height=7, exportselection=0, font=(
                font_name, 15, "bold"))  # リストボックスに追加
            self.listkanabox.grid(column=2, row=2, sticky=W + E + N + S)

            button_1 = ttk.Button(self.dialog, text="検索", padding=[
                                  330, 20, 330, 20], style='office.TButton')
            button_1.bind('<Button-1>', func=self.selectCY)
            button_1.grid(column=0, row=3, sticky=W + E + N + S, columnspan=3)

            font4 = font.Font(family=font_name, size=15, weight='bold')
            self.label4 = Label(self.dialog, text="名前一覧", font=font4)
            self.label4.grid(column=1, row=4, sticky=W + E + N + S)

            self.selectbox = Listbox(self.dialog, listvariable=selectyearname,
                                     height=15, exportselection=0, font=(font_name, 15, "bold"))
            self.selectbox.grid(
                column=0, row=5, columnspan=3, sticky=W + E + N + S)

            button_2 = ttk.Button(self.dialog, text="確定",  padding=[
                                  330, 20, 330, 20], style='office.TButton')
            button_2.bind('<Button-1>', func=self.selectNAME)
            button_2.grid(column=0, row=6, sticky=W + E + N + S, columnspan=3)

        elif Bool_value == False:
            self.reload_modules()
            messagebox.showerror('ファイル参照エラー', '生徒の名簿ファイルが正しくありません')

    def selectCY(self, event):
        global choose_list, select_course, select_year
        self.reload_modules()
        # 選択されている数値インデックスを含むリストを取得
        itemIdxList1 = self.listcoursebox.curselection()
        itemIdxList2 = self.listyearbox.curselection()
        itemIdxList3 = self.listkanabox.curselection()
        if self.selectbox.size() >= 1:
            self.selectbox.delete(0, self.selectbox.size())

        if len(itemIdxList1) == 1:
            if len(itemIdxList1) == 1:
                select_course = self.course[itemIdxList1[0]]
                select_year = self.year[itemIdxList2[0]]
                kana_num = itemIdxList3[0]
                Bool_value, list_value, yearName = Sort_Students.load_studentlist(
                    JsonReader.Read_json("./settings.json"))
                student_list_keys = Sort_Students.setlist_keys(list_value)
                choose_list = Sort_Students.choose_CYname(
                    student_list_keys, select_course, select_year, kana_num)
                # 末尾に選択された要素を追加する
                for i in choose_list:
                    self.selectbox.insert("end", i)

    def selectNAME(self, event):
        global choose_list, Select_Student, Select_Number, select_course, select_year
        self.reload_modules()
        SeatNumber.New_day_book()
        # 選択されている数値インデックスを含むリストを取得
        itemIdxList = self.selectbox.curselection()
        if len(itemIdxList) == 1:
            select_name = choose_list[itemIdxList[0]]
            self.dialog.destroy()
            Select_Student = select_name
            (buttons[Select_Number - 1])['style'] = 'MyWidget2.TButton'
            SeatNumber.append_time_in(
                Select_Student, Select_Number, select_year, select_course)

    def OpenDialog(self):  # 席を開けるときのダイアログ
        # ウィジェットの作成、配置
        global Select_Number, Select_Student, time

        self.reload_modules()
        Select_Student = SeatNumber.give_name(Select_Number)

        check_Seat = messagebox.askyesno(
            title=f"{Select_Number}番の席",
            message=f"{Select_Student} さん、この席を空けますか?                     ")

        if check_Seat:
            self.reload_modules()
            (buttons[Select_Number - 1])['style'] = 'MyWidget.TButton'
            time = SeatNumber.leave_seat_time(Select_Student, Select_Number)
            messagebox.showinfo(
                title=f"{Select_Number}番の席",
                message=f"{Select_Student} さん、お疲れ様でした。                     ",
                detail=f"今日の勉強時間は {time} です。                     ")


# スプラッシュスクリーン作成
splash = Tk()
splash.overrideredirect(1)  # スプラッシュ画面のタイトルバー非表示

photo = PhotoImage(file=settings.pythonLOGO)  # 表示させたい画像ファイル指定

splash_width = 800
splash_height = 600
splash_x = int(int(splash.winfo_screenwidth()/2) - int(splash_width/2))
splash_y = int(int(splash.winfo_screenheight()/2) - int(splash_height/2))
splash.geometry(f"{splash_width}x{splash_height}+{splash_x}+{splash_y}")

canvas = Canvas(splash)  # 画像のサイズに合わせて幅と高さ調整
canvas.pack(expand=True, fill=BOTH)
canvas.update()
canvas_width = canvas.winfo_width()
canvas_height = canvas.winfo_height()
canvas.create_image(canvas_width / 2, canvas_height /
                    2, image=photo)


def main():
    splash.destroy()
    root = Tk()
    root.overrideredirect(1)
    root.title('座席表')
    root.geometry("1920x1080")
    root.state("zoomed")
    root.iconbitmap(settings.pythonLOGOICO)
    Seat(root)

    s = ttk.Style()
    s.theme_use('alt')

    root.mainloop()


if __name__ == '__main__':
    splash.after(5000, main)
    splash.mainloop()
