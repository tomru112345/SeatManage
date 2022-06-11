import tkinter
import tkinter.filedialog
import tkinter.ttk
import tkinter.font
import tkinter.messagebox
import settings


class App(tkinter.Frame):
    def __init__(self, master):
        """初期化"""
        super().__init__(master)
        self.master.state("zoomed")
        self.master.title("座席表マネージャ")
        self.master.style = tkinter.ttk.Style()
        self.setting_style()

    def seating_chart(self):
        # Frame
        top_frame = tkinter.ttk.Frame(self.master, style="TFrame")

        # ノートブック
        notebook = tkinter.ttk.Notebook(self.master, style="TNotebook")

        # タブの作成
        tab_list = []
        for _ in range(len(settings.room_list)):
            tab_list.append(tkinter.ttk.Frame(notebook, style="TFrame"))

        # 一番上のラベルの設定
        label = []
        label.append(tkinter.ttk.Label(top_frame, text="座席表", style="header.TLabel"))

        if len(tab_list) > 0:
            for i in range(len(settings.room_list)):
                # notebookにタブを追加
                notebook.add(tab_list[i], text=settings.room_list[i])
                # タブの中のラベルの設定
                label.append(tkinter.ttk.Label(
                    tab_list[i], text=settings.room_list[i], style="h1.TLabel"))

            notebook.pack(expand=True, fill=tkinter.BOTH,
                          padx=10, pady=10, side=tkinter.BOTTOM)
        else:
            # error_frame
            error_frame = tkinter.ttk.Frame(self.master, style="TFrame")
            label.append(tkinter.ttk.Label(
                error_frame, text="座席表のデータが読み込めません", style="error.h1.TLabel"))
            # ウィジェットの配置
            error_frame.pack(expand=True, fill=tkinter.BOTH, padx=10,
                             pady=10, side=tkinter.BOTTOM)

        # ウィジェットの配置
        top_frame.pack(fill=tkinter.BOTH, padx=10, pady=10)
        for tmp in label:
            tmp.pack()

    def setting_style(self):
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


def main():
    root = tkinter.Tk()
    app = App(master=root)  # Inherit
    app.seating_chart()
    app.mainloop()


main()
