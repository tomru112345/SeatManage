import tkinter
import os
import tkinter.filedialog
import tkinter.ttk
import tkinter.font
import tkinter.messagebox
import settings
import logmanager
import reader
import importlib
import sys


class Seat(tkinter.ttk.Frame):  # リストボックスのクラス

    def __init__(self, master=None):
        """初期化"""
        super().__init__(master)
        self.pingfile = settings.pingfile

        self.Select_Student = ""
        self.Select_Number = 0
        self.select_course = ""
        self.select_year = ""
        self.time = ""

        self.READ_DATA = reader.Reader()
        self.font_name = settings.font_name

        self.buttons = []
        self.create_style()
        self.create_widgets()
        pass

    def reload_modules(self):
        """リロード"""
        importlib.reload(settings)
        importlib.reload(logmanager)

    def create_style(self):
        """ボタン、ラベルのスタイルを変更."""
        style = tkinter.ttk.Style()

        style_kind = ['winnative', 'clam', 'alt',
                      'default', 'classic', 'vista', 'xpnative']
        i = 2

        style.theme_use(style_kind[i])
        # ボタンのスタイルを上書き
        style.configure(
            'MyWidget.TButton',
            font=(
                self.font_name,
                20
            ),
            background='#32CD32'
        )

        style2 = tkinter.ttk.Style()
        style2.theme_use(style_kind[i])
        # ボタンのスタイルを上書き
        style2.configure('office.TButton', font=(
            self.font_name, 20), background='#D3D3D3')

        style3 = tkinter.ttk.Style()
        style3.theme_use(style_kind[i])
        # ボタンのスタイルを上書き
        style3.configure('MyWidget2.TButton', font=(
            self.font_name, 20), background='#DC143C')

        style4 = tkinter.ttk.Style()
        style4.theme_use(style_kind[i])
        # ボタンのスタイルを上書き
        style4.configure('office2.TButton', font=(
            self.font_name, 10), background='#D3D3D3')

    def create_widgets(self):
        """席ボタンウィジェットの作成."""
        font0 = tkinter.font.Font(
            family=self.font_name, size=20, weight='bold')
        self.label0 = tkinter.ttk.Label(self, text="""
        自習室の希望する席を選んでください。
        * 緑 : 席が空いてます
        * 赤 : 席を使っています
        """, font=font0, anchor='e', justify='left')
        self.label0.grid(column=0, row=0, columnspan=5)

        # レイアウトの作成
        btn_num = 0
        for y, row in enumerate(self.READ_DATA.LAYOUT, 1):
            for x, char in enumerate(row):
                if char != "x":
                    No_Vacant_Seat = logmanager.LogManager().Open()
                    if (btn_num + 1) in No_Vacant_Seat:
                        self.buttons.append(tkinter.ttk.Button(self, text=str(
                            btn_num + 1), style='MyWidget2.TButton'))
                    else:
                        self.buttons.append(tkinter.ttk.Button(self, text=str(
                            btn_num + 1), style='MyWidget.TButton'))
                    self.buttons[btn_num].grid(
                        column=x, row=y, sticky=(tkinter.N, tkinter.S, tkinter.E, tkinter.W))
                    self.buttons[btn_num].bind(
                        '<Button-1>', func=self.click_option)
                    btn_num += 1

        self.grid(column=0, row=0, sticky=(
            tkinter.N, tkinter.S, tkinter.E, tkinter.W))

        # 横の引き伸ばし設定
        for i in range(self.READ_DATA.LENGTH):
            self.columnconfigure(i, weight=1)

        # 縦の引き伸ばし設定。0番目の結果表示欄だけ、元の大きさのまま
        self.rowconfigure(0, weight=0)
        for i in range(self.READ_DATA.HEIGHT):
            self.rowconfigure(i + 1, weight=1)

        # ウィンドウ自体の引き伸ばし設定
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)

        # menubarの大元（コンテナ）の作成と設置
        menubar = tkinter.Menu(self)
        file_menu = tkinter.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="設定", menu=file_menu)

        # 生徒名簿の設定
        file_menu.add_command(
            label="生徒名簿", command=self.onOpenSettingStudentfile, accelerator="Ctrl+O")
        self.master.config(menu=menubar)
        self.bind_all("<Control-o>", self.onOpenSettingStudentfile)

        # 座席表の設定
        file_menu.add_command(
            label="座席表", command=self.onOpenSettingSeat, accelerator="Ctrl+T")
        self.master.config(menu=menubar)
        self.bind_all("<Control-t>", self.onOpenSettingSeat)

        file_menu.add_command(
            label="終了", command=self.ExitApp, accelerator="Ctrl+F")
        self.master.config(menu=menubar)
        self.bind_all("<Control-f>", self.ExitApp)

        # ライセンス表示
        License_menu = tkinter.Menu(menubar, tearoff=0)
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

        # AppID の設定
        License_menu.add_command(
            label="App ID", command=self.onOpenSettingAppID, accelerator="Ctrl+k")
        self.master.config(menu=menubar)
        self.bind_all("<Control-k>", self.onOpenSettingAppID)

    def ExitApp(self, event=None):
        check_Fin = tkinter.messagebox.askyesno(
            title="アプリケーション終了",
            message="アプリケーションを終了しますか？")

        if check_Fin:
            self.reload_modules()
            sys.exit()

    def onOpenSettingID(self, event=None):
        """学年 ID"""
        self.reload_modules()
        self.dialog = tkinter.Toplevel(self)
        self.dialog.iconbitmap(settings.pythonLOGOICO)
        self.dialog.title("学年 ID")
        self.dialog.geometry('400x300')
        self.dialog.resizable(width=False, height=False)
        # Bool_value, self.student_list, yearName = SortStudents.load_studentlist(self.path)
        Bool_value = self.READ_DATA.ExcelRead()

        # 列の識別名を指定
        column = ('ID', 'SchoolYear')

        # Treeviewの生成
        tree = tkinter.ttk.Treeview(self.dialog, columns=column)

        # 列の設定
        tree.column('#0', width=0, stretch='no')
        tree.column('ID', anchor='center', width=80)
        tree.column('SchoolYear', anchor='center', width=80)

        # 列の見出し設定
        tree.heading('#0', text='')
        tree.heading('ID', text='ID', anchor='center')
        tree.heading('SchoolYear', text='学年', anchor='center')

        if Bool_value:
            t = 0
            for i in self.READ_DATA.DicYear.keys():
                # for i in self.School_year.keys():
                # レコードの追加
                # tree.insert(parent='', index='end', iid= t, values=(self.School_year[i], i), tags = t)
                tree.insert(parent='', index='end', iid=t, values=(
                    self.READ_DATA.DicYear[i], i), tags=t)
                if t & 1:
                    tree.tag_configure(t, background="#CCFFFF")

                t += 1

        # ウィジェットの配置
        tree.pack()

    def onOpenSettingSeat(self, event=None):
        """座席表の設定"""
        self.reload_modules()
        self.dialog = tkinter.Toplevel(self)
        self.dialog.iconbitmap(settings.pythonLOGOICO)
        self.dialog.title("座席表の設定")
        self.dialog.geometry('400x300')
        self.dialog.resizable(width=False, height=False)

    def onOpenSettingAppID(self, event=None):
        """学年 ID"""
        self.reload_modules()
        self.dialog = tkinter.Toplevel(self)
        self.dialog.iconbitmap(settings.pythonLOGOICO)
        self.dialog.title("App ID")
        window_width = 350
        window_height = 700
        x = int(
            int(self.dialog.winfo_screenwidth() / 2) - int(window_width / 2)
        )
        y = int(
            int(self.dialog.winfo_screenheight() / 2) - int(window_height / 2)
        )
        self.dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.dialog.resizable(width=False, height=False)
        self.dialog.grab_set()

    def onOpenLicense(self, event=None):
        """ライセンス"""
        tkinter.messagebox.showinfo(
            title="ライセンス",
            message="ライセンス",
            detail=settings.License)

    def onOpenSettingStudentfile(self, event=None):
        """生徒名簿の設定"""
        # リロード
        self.reload_modules()

        self.dialog = tkinter.Toplevel(self)
        self.dialog.iconbitmap(settings.pythonLOGOICO)
        self.dialog.title("生徒名簿の設定")
        window_width = 480
        window_height = 150
        x = int(
            int(self.dialog.winfo_screenwidth() / 2) - int(window_width / 2)
        )
        y = int(
            int(self.dialog.winfo_screenheight() / 2) - int(window_height / 2)
        )
        self.dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.dialog.resizable(width=False, height=False)
        self.dialog.grab_set()

        font1 = tkinter.font.Font(
            family=self.font_name, size=10, weight='bold')
        self.label1 = tkinter.Label(
            self.dialog,
            text="生徒名簿の Excel ファイルを指定してください。",
            font=font1,
            anchor='e',
            justify='left'
        )

        self.label1.grid(
            column=0,
            row=0,
            columnspan=4,
            sticky=tkinter.EW,
            padx=5
        )

        button_file = tkinter.ttk.Button(
            self.dialog,
            text="開く",
            style="office2.TButton"
        )

        button_file.bind(
            '<Button-1>',
            func=self.file_dialog
        )

        button_file.grid(
            column=4,
            row=0,
            sticky=tkinter.E
        )

        self.label_name = tkinter.Label(
            self.dialog,
            text="選択ファイル名:",
            font=font1,
            justify='left'
        )
        self.label_name.grid(column=0, row=1)

        self.file_name = tkinter.StringVar()
        # self.file_name.set(self.StudentsListPath)
        self.file_name.set(self.READ_DATA.EXCEL_PATH)
        self.label2 = tkinter.Label(
            self.dialog,
            textvariable=self.file_name,
            font=font1,
            justify='left'
        )
        self.label2.grid(column=1, row=1, columnspan=4)

        button_fin = tkinter.ttk.Button(
            self.dialog,
            text="決定",
            style="office2.TButton"
        )
        button_fin.bind(
            '<Button-1>',
            func=self.select_filename
        )
        button_fin.grid(column=4, row=2)

    def file_dialog(self, event):
        """ファイルの選択オプション"""
        fTyp = [("Excel", "xlsx")]
        iDir = os.path.abspath(os.path.dirname(__file__))
        file_name = tkinter.filedialog.askopenfilename(
            filetypes=fTyp,
            initialdir=iDir
        )
        if len(file_name) == 0:
            # self.file_name.set(self.StudentsListPath)
            self.file_name.set(self.READ_DATA.EXCEL_PATH)
        else:
            self.file_name.set(file_name)

    def select_filename(self, event):
        """生徒名簿ファイル選択結果の更新を行う関数"""
        # filename = self.file_name.get()
        self.READ_DATA.JsonWrite()
        self.dialog.destroy()

    def click_option(self, event):
        global Select_Number
        Select_Number = int(event.widget.cget("text"))
        if (self.buttons[Select_Number - 1])['style'] == 'MyWidget.TButton':
            self.OpenListbox()
        elif (self.buttons[Select_Number - 1])['style'] == 'MyWidget2.TButton':
            self.OpenDialog()

    def OpenListbox(self):  # 席を取るときのダイアログ
        """リストウィジェットの作成"""
        self.reload_modules()
        Bool_value = self.READ_DATA.ExcelRead()

        if Bool_value:
            self.READ_DATA.ExcelRead()
            # School_year_ID = SortStudents.setlist_ID(self.student_list)
            # self.READ_DATA.set_ListCourse()
            # course = SortStudents.setlist_course(self.student_list, settings.course)

            self.reload_modules()
            self.dialog = tkinter.Toplevel(self)
            self.dialog.iconbitmap(settings.pythonLOGOICO)
            self.dialog.title("生徒リスト")
            window_width = 960
            window_height = 960
            x = int(
                int(self.dialog.winfo_screenwidth() / 2) - int(window_width / 2)
            )
            self.dialog.geometry(f"{window_width}x{window_height}+{x}+{0}")
            self.dialog.grab_set()

            # 横の引き伸ばし設定
            for i in range(3):
                self.dialog.columnconfigure(i, weight=1)

            # 縦の引き伸ばし設定。0番目の結果表示欄だけ、元の大きさのまま
            self.rowconfigure(0, weight=0)
            for i in range(7):
                self.dialog.rowconfigure(i, weight=1)

            font1 = tkinter.font.Font(
                family=self.font_name, size=15, weight='bold')
            self.label1 = tkinter.Label(self.dialog, text=settings.text_set,
                                        font=font1, anchor='e', justify='left')
            self.label1.grid(column=0, row=0, columnspan=3)

            font2 = tkinter.font.Font(
                family=self.font_name, size=15, weight='bold')
            self.label2 = tkinter.Label(self.dialog, text="学年", font=font2)
            self.label2.grid(column=0, row=1,
                             sticky=tkinter.W + tkinter.E + tkinter.N + tkinter.S)

            font3 = tkinter.font.Font(
                family=self.font_name, size=15, weight='bold')
            self.label3 = tkinter.Label(self.dialog, text="コース", font=font3)
            self.label3.grid(column=1, row=1,
                             sticky=tkinter.W + tkinter.E + tkinter.N + tkinter.S)

            self.label3 = tkinter.Label(self.dialog, text="カナ行", font=font3)
            self.label3.grid(column=2, row=1,
                             sticky=tkinter.W + tkinter.E + tkinter.N + tkinter.S)

            # self.year = self.yearName  # 学年リスト
            self.year = self.READ_DATA.yearName
            self.yearname = tkinter.StringVar(value=self.year)  # 文字列を保持させる
            selectyearname = tkinter.StringVar()  # 文字列を保持させる

            self.listyearbox = tkinter.Listbox(self.dialog, listvariable=self.yearname, height=7, exportselection=0, font=(
                self.font_name, 15, "bold"))  # リストボックスに追加
            self.listyearbox.grid(
                column=0, row=2, sticky=tkinter.W + tkinter.E + tkinter.N + tkinter.S)

            # self.course = course  # コースリスト
            self.course = self.READ_DATA.COURSE
            coursename = tkinter.StringVar(value=self.course)  # 文字列を保持させる

            self.listcoursebox = tkinter.Listbox(self.dialog, listvariable=coursename, height=7, exportselection=0, font=(
                self.font_name, 15, "bold"))  # リストボックスに追加
            self.listcoursebox.grid(
                column=1, row=2, sticky=tkinter.W + tkinter.E + tkinter.N + tkinter.S)

            self.kana = ["ア行", "カ行", "サ行", "タ行", "ナ行",
                         "ハ行", "マ行", "ヤ行", "ラ行", "ワ行"]  # コースリスト
            kana = tkinter.StringVar(value=self.kana)  # 文字列を保持させる

            self.listkanabox = tkinter.Listbox(self.dialog, listvariable=kana, height=7, exportselection=0, font=(
                self.font_name, 15, "bold"))  # リストボックスに追加
            self.listkanabox.grid(
                column=2, row=2, sticky=tkinter.W + tkinter.E + tkinter.N + tkinter.S)

            button_1 = tkinter.ttk.Button(self.dialog, text="検索", padding=[
                330, 20, 330, 20], style='office.TButton')
            button_1.bind('<Button-1>', func=self.selectCY)
            button_1.grid(column=0, row=3,
                          sticky=tkinter.W + tkinter.E + tkinter.N + tkinter.S, columnspan=3)

            font4 = tkinter.font.Font(
                family=self.font_name, size=15, weight='bold')
            self.label4 = tkinter.Label(self.dialog, text="名前一覧", font=font4)
            self.label4.grid(column=1, row=4,
                             sticky=tkinter.W + tkinter.E + tkinter.N + tkinter.S)

            self.selectbox = tkinter.Listbox(self.dialog, listvariable=selectyearname,
                                             height=15, exportselection=0, font=(self.font_name, 15, "bold"))
            self.selectbox.grid(
                column=0, row=5, columnspan=3, sticky=tkinter.W + tkinter.E + tkinter.N + tkinter.S)

            button_2 = tkinter.ttk.Button(self.dialog, text="確定", padding=[
                330, 20, 330, 20], style='office.TButton')
            button_2.bind('<Button-1>', func=self.selectNAME)
            button_2.grid(column=0, row=6,
                          sticky=tkinter.W + tkinter.E + tkinter.N + tkinter.S, columnspan=3)

        elif not Bool_value:
            self.reload_modules()
            tkinter.messagebox.showerror('ファイル参照エラー', '生徒の名簿ファイルが正しくありません')

    def selectCY(self, event):
        global choose_list, select_course, select_year
        self.reload_modules()
        # 選択されている数値インデックスを含むリストを取得
        itemIdxLists = []
        itemIdxLists.append(self.listcoursebox.curselection())
        itemIdxLists.append(self.listyearbox.curselection())
        itemIdxLists.append(self.listkanabox.curselection())

        if self.selectbox.size() >= 1:
            self.selectbox.delete(0, self.selectbox.size())

        if len(itemIdxLists[0]) == 1:
            select_course = self.course[itemIdxLists[0][0]]
            select_year = self.year[itemIdxLists[1][0]]
            kana_num = itemIdxLists[2][0]
            self.READ_DATA.ExcelRead()
            # self.READ_DATA.set_ListKeys()
            self.READ_DATA.ExcelRead()
            choose_list = self.READ_DATA.choose_CYname(
                select_course, select_year, kana_num)
            # 末尾に選択された要素を追加する
            for i in choose_list:
                self.selectbox.insert("end", i)

    def selectNAME(self, event):
        global choose_list, Select_Student, Select_Number, select_course, select_year
        self.reload_modules()
        logmanager.LogManager().Create()
        # 選択されている数値インデックスを含むリストを取得
        itemIdxList = self.selectbox.curselection()
        if len(itemIdxList) == 1:
            select_name = choose_list[itemIdxList[0]]
            self.dialog.destroy()
            Select_Student = select_name
            (self.buttons[Select_Number - 1])['style'] = 'MyWidget2.TButton'
            logmanager.LogManager().LOG_Append(
                Select_Student, Select_Number, select_year, select_course)

    def OpenDialog(self):  # 席を開けるときのダイアログ
        # ウィジェットの作成、配置
        global Select_Number, Select_Student, time

        self.reload_modules()
        Select_Student = logmanager.LogManager().give_name(Select_Number)

        check_Seat = tkinter.messagebox.askyesno(
            title=f"{Select_Number}番の席",
            message=f"{Select_Student} さん、この席を空けますか?                     "
        )

        if check_Seat:
            self.reload_modules()
            (self.buttons[Select_Number - 1])['style'] = 'MyWidget.TButton'
            time = logmanager.LogManager().LOG_Leave(Select_Student, Select_Number)
            tkinter.messagebox.showinfo(
                title=f"{Select_Number}番の席",
                message=f"{Select_Student} さん、お疲れ様でした。                     ",
                detail=f"今日の勉強時間は {time} です。                     ")
