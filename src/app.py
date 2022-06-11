import tkinter
import tkinter.filedialog
import tkinter.ttk
import tkinter.font
import tkinter.messagebox
import settings


class App(tkinter.Frame):
    def __init__(self, master: tkinter.Tk):
        """初期化"""
        super().__init__(master)
        # マスターの設定
        self.master.geometry("1920x1080")
        self.master.state("zoomed")
        self.master.title("座席表マネージャ")

        # スタイルの設定
        self.master.style = tkinter.ttk.Style()
        self.set_style()

        # タブのリスト
        self.master.tabs = []

        # ラベルのリスト
        self.master.labels = []

        # ボタンのリスト
        self.master.buttons = []

    def seating_chart(self) -> None:
        # Frame
        # top_frame = tkinter.ttk.Frame(self.master, style="TFrame")

        # ノートブック
        notebook = tkinter.ttk.Notebook(self.master, style="TNotebook")

        # タブの作成
        for _ in range(len(settings.room_list)):
            self.master.tabs.append(
                tkinter.ttk.Frame(notebook, style="TFrame"))

        # 一番上のラベルの設定
        # self.master.labels.append(tkinter.ttk.Label(
        #     top_frame, text="座席表", style="header.TLabel"))

        if len(self.master.tabs) > 0:
            for i in range(len(settings.room_list)):
                # notebookにタブを追加
                notebook.add(self.master.tabs[i], text=settings.room_list[i])

                # ラベル用のフレームの作成
                labels_frame = tkinter.ttk.Frame(
                    self.master.tabs[i], style="TFrame")
                # ボタン用のフレームの作成
                bottoms_frame = tkinter.ttk.Frame(
                    self.master.tabs[i], style="TFrame", relief="groove")

                # タブの中のラベルの設定
                self.master.labels.append(tkinter.ttk.Label(
                    labels_frame, text=settings.room_list[i], style="h1.TLabel"))

                # ボタンを作成
                for tmp in range(3):
                    self.master.buttons.append(
                        tkinter.ttk.Button(bottoms_frame, text=str(tmp + 1), style="TButton"))

                # ラベル用のフレームの作成
                labels_frame.pack(fill=tkinter.BOTH, padx=10,
                                  pady=10, side=tkinter.TOP)
                # ボタン用フレームの配置
                bottoms_frame.pack(expand=True, fill=tkinter.BOTH, padx=10,
                                   pady=10, side=tkinter.BOTTOM)

            notebook.pack(expand=True, fill=tkinter.BOTH,
                          padx=10, pady=10, side=tkinter.BOTTOM)
        else:
            # error_frame
            error_frame = tkinter.ttk.Frame(self.master, style="TFrame")
            self.master.labels.append(tkinter.ttk.Label(
                error_frame, text="座席表のデータが読み込めません", style="error.h1.TLabel"))
            # ウィジェットの配置
            error_frame.pack(expand=True, fill=tkinter.BOTH, padx=10,
                             pady=10, side=tkinter.BOTTOM)

        # ウィジェットの配置
        # top_frame.pack(fill=tkinter.BOTH, padx=10, pady=10)
        for tmp in self.master.labels:
            tmp.pack()
        for tmp in self.master.buttons:
            tmp.pack()

    def set_style(self) -> None:
        # スタイルの設定
        self.master.style.theme_use('clam')

        # Syle (TNotebook)
        self.master.style.configure(
            "TNotebook",
            tabposition=tkinter.SW,
        )
        # Syle (TNotebook.Tab)
        self.master.style.configure(
            "TNotebook.Tab",
            font=("Meiryo", 15, "bold"),
            background="white",
            foreground="black",
            justify=tkinter.CENTER
        )
        # Style.map (TNotebook.Tab)
        self.master.style.map(
            "TNotebook.Tab",
            foreground=[
                ('active', 'white'),
                ('disabled', 'gray'),
                ('selected', 'blue'),
            ],
            background=[
                ('active', 'darkorange'),
                ('disabled', 'black'),
                ('selected', 'white'),
            ],
        )

        # Syle (TLabel)
        self.master.style.configure(
            "header.TLabel",
            font=("Meiryo", 30, "bold"),
            foreground="black",
            background="white",
            justify=tkinter.CENTER
        )

        self.master.style.configure(
            "h1.TLabel",
            font=("Meiryo", 30, "bold"),
            foreground="black",
            background="white",
            side=tkinter.TOP
        )

        self.master.style.configure(
            "error.h1.TLabel",
            font=("Meiryo", 30, "bold"),
            foreground="red",
            background="white",
            justify=tkinter.CENTER
        )

        self.master.style.configure(
            "TFrame",
            background="white"
        )

        # Style
        self.master.style.configure(
            "TButton",
            font=("Meiryo", 30, "bold"),
            width=10,
            background="medium spring green",
            foreground="black",
            justify=tkinter.CENTER
        )

        self.master.style.map(
            "TButton",
            foreground=[
                ('active', 'white'),
                ('disabled', 'gray'),
                ('selected', 'blue'),
            ],
            background=[
                ('active', 'SteelBlue1'),
                ('disabled', 'black'),
                ('selected', 'orange red'),
            ],
        )


def main():
    root = tkinter.Tk()
    app = App(master=root)  # Inherit
    app.seating_chart()
    app.mainloop()


main()
