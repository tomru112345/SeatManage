# -*- coding:utf-8 -*-
from tkinter import *
from tkinter import font
from tkinter import ttk
import Sort_Students
import settings
import SeatNumber

height = settings.height
length = settings.length

pingfile = settings.pingfile

Select_Student = ""
Select_Number = 0

# 2次元配列のとおりに、gridでレイアウトを作成する
LAYOUT = [
    ['1', '2', '3', '4', '5'],
    ['6', '7', '8', '9', '10'],
    ['11', '12', '13', '14', '15'],
    ['16', '17', '18', '19', 'x'],
    ['20', '21', '22', '23', '24'],
    ['25', '26', '27', '28', '29'],
    ['30', '31', '32', '33', '34']
]

buttons = []

class Seat(ttk.Frame): # リストボックスのクラス
    
    def __init__(self, master=None):
        super().__init__(master)
        self.create_style()
        self.create_widgets()

    def create_style(self):
        """ボタン、ラベルのスタイルを変更."""
        style = ttk.Style()
        style.theme_use('alt')
        # ボタンのスタイルを上書き
        style.configure('MyWidget.TButton', font=('Helvetica', 20), background='#00DD00')

        style2 = ttk.Style()
        style2.theme_use('alt')
        # ボタンのスタイルを上書き
        style2.configure('office.TButton', font=('Helvetica', 20), background='#FFFFFF')

        style3 = ttk.Style()
        style3.theme_use('alt')
        # ボタンのスタイルを上書き
        style3.configure('MyWidget2.TButton', font=('Helvetica', 20), background='#FF0000')

    def create_widgets(self):
        """ウィジェットの作成."""
        font0 = font.Font(size=20, weight='bold')
        self.label0 = Label(self, text="""
        自習室の希望する席を選んでください。
        * 緑 : 席が空いてます
        * 赤 : 席が使われています
        """, font = font0, anchor='e', justify='left')
        self.label0.grid(column=0, row=0, columnspan=5)
        # レイアウトの作成
        for y, row in enumerate(LAYOUT, 1):
            for x, char in enumerate(row):
                if char != "x":
                    No_Vacant_Seat = SeatNumber.Open_book()
                    if int(char) in No_Vacant_Seat:
                        index = int(char) - 1
                        buttons.append(ttk.Button(self, text=char, style= 'MyWidget2.TButton'))
                        buttons[index].grid(column=x, row=y, sticky=(N, S, E, W))
                        buttons[index].bind('<Button-1>', func = self.click_option)
                    else: 
                        index = int(char) - 1
                        buttons.append(ttk.Button(self, text=char, style= 'MyWidget.TButton'))
                        buttons[index].grid(column=x, row=y, sticky=(N, S, E, W))
                        buttons[index].bind('<Button-1>', func = self.click_option)
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

    def click_option(self, event):
        global Select_Number
        Select_Number = int(event.widget.cget("text"))
        if (buttons[Select_Number - 1])['style'] == 'MyWidget.TButton':
            self.OpenListbox()
        elif (buttons[Select_Number - 1])['style'] == 'MyWidget2.TButton':
            self.OpenDialog()
            
        
    def OpenListbox(self): # 席を取るときのダイアログ
        #ウィジェットの作成、配置
        self.dialog = Toplevel(self)
        self.dialog.title("生徒リスト")
        self.dialog.geometry("860x850")
        self.dialog.resizable(width=False, height=False)
        self.dialog.grab_set()

        font1 = font.Font(size=20, weight='bold')
        self.label1 = Label(self.dialog, text="""
        [1] 学年, [2] コースを選んで A のボタンを押してください。
        [3] 名前 の自分の名前を選んで B のボタンを押して下さい。
        """, font = font1, anchor='e', justify='left')
        self.label1.grid(column=0, row=0, columnspan=2)

        font2 = font.Font(size=20, weight='bold')
        self.label2 = Label(self.dialog, text="[1] 学年", font = font2)
        self.label2.grid(column=0, row=1)

        font3 = font.Font(size=20, weight='bold')
        self.label3 = Label(self.dialog, text="[2] コース", font = font3)
        self.label3.grid(column=1, row=1)

        self.year = settings.School_year # 学年リスト
        yearname = StringVar(value=self.year) # 文字列を保持させる
        selectyearname = StringVar() # 文字列を保持させる

        self.listyearbox  =  Listbox(self.dialog, listvariable=yearname, height=8, exportselection=0, font=("",20)) # リストボックスに追加
        self.listyearbox.grid(column=0, row=2)

        self.course = Sort_Students.course # コースリスト
        coursename = StringVar(value=self.course) # 文字列を保持させる

        self.listcoursebox  =  Listbox(self.dialog, listvariable=coursename, height=8, exportselection=0, font=("",20)) # リストボックスに追加
        self.listcoursebox.grid(column=1, row=2)

        button_1 = ttk.Button(self.dialog, text = "A" ,command=self.selectCY, padding=[330,20,330,20], style="office.TButton")
        button_1.grid(column=0, row=3, sticky = N,columnspan= 2)

        font4 = font.Font(size=20, weight='bold')
        self.label4 = Label(self.dialog, text="[3] 名前", font = font4)
        self.label4.grid(column=0, row=4)

        self.selectbox = Listbox(self.dialog, listvariable=selectyearname, height=8, exportselection=0, font=("",20))
        self.selectbox.grid(column=0, row=5)

        button_2 = ttk.Button(self.dialog, text = "B" ,command=self.selectNAME,  padding=[330,20,330,20], style="office.TButton")
        button_2.grid(column=0, row=6, sticky = N,columnspan= 2)
        

    def selectCY(self):
        global choose_list
        # 選択されている数値インデックスを含むリストを取得
        itemIdxList1 =  self.listcoursebox.curselection()
        itemIdxList2 =  self.listyearbox.curselection()
        if self.selectbox.size() >= 1:
            self.selectbox.delete(0, self.selectbox.size())

        if len(itemIdxList1) == 1:
            if len(itemIdxList1) == 1:
                select_course = self.course[itemIdxList1[0]]
                select_year = self.year[itemIdxList2[0]]
                choose_list = Sort_Students.choose_CYname(select_course, select_year)
                # 末尾に選択された要素を追加する
                for i in choose_list:
                    self.selectbox.insert("end", i)

    def selectNAME(self):
        global choose_list, Select_Student, Select_Number
        SeatNumber.New_day_book()
        # 選択されている数値インデックスを含むリストを取得
        itemIdxList =  self.selectbox.curselection()
        if len(itemIdxList) == 1:
            select_name = choose_list[itemIdxList[0]]
            #self.destroy()
            self.dialog.destroy()
            Select_Student = select_name
            (buttons[Select_Number - 1])['style'] = 'MyWidget2.TButton'
            SeatNumber.append_time_in(Select_Student, Select_Number)

    def OpenDialog(self): # 席を開けるときのダイアログ
        #ウィジェットの作成、配置
        self.dialog = Toplevel(self)
        self.dialog.title("座席コンフィグ")
        self.dialog.geometry("440x200")
        self.dialog.resizable(width=False, height=False)
        self.dialog.grab_set()

        font1 = font.Font(size=20, weight='bold')
        self.label1 = Label(self.dialog, text="""
        この席を開けますか?
        """, font = font1, anchor='w')
        self.label1.grid(column=0, row=0, columnspan=3)

        button_yes = ttk.Button(self.dialog, text = "はい" ,command=self.selectYes, style="office.TButton")
        button_yes.grid(column=0, row=1, sticky=(N, S, E, W))

        button_no = ttk.Button(self.dialog, text = "いいえ" ,command=self.selectNo, style="office.TButton")
        button_no.grid(column=2, row=1, sticky=(N, S, E, W))

    def selectYes(self):
        global Select_Number
        self.dialog.destroy()
        (buttons[Select_Number - 1])['style'] = 'MyWidget.TButton'
        SeatNumber.leave_seat_time(Select_Number)
    
    def selectNo(self):
        self.dialog.destroy()


def main():
    root = Tk()
    root.title('座席表')
    root.geometry("1920x1080")
    root.call('wm', 'iconphoto', root._w, PhotoImage(file=pingfile))
    Seat(root)
    root.mainloop()


if __name__ == '__main__':
    main()