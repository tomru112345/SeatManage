import tkinter
import tkinter.filedialog
import tkinter.ttk
import tkinter.font
import tkinter.messagebox


class App(tkinter.Frame):
    def __init__(self, master):
        """初期化"""
        super().__init__(master)
        # self.pack()
        self.master.geometry("900x900")
        self.master.title("座席表マネージャ")

        # Frame
        top_frame = tkinter.ttk.Frame(self.master, style="TFrame")

        # ノートブック
        notebook = tkinter.ttk.Notebook(self.master, style="TNotebook")

        style = tkinter.ttk.Style()

        style.theme_use('clam')

        # スタイルの設定
        # Syle (TNotebook)
        style.configure(
            "TNotebook",
            tabposition=tkinter.SW,
        )
        # Syle (TNotebook.Tab)
        style.configure(
            "TNotebook.Tab",
            font=("Meiryo", 15, "bold"),
            background="white",
            foreground="black",
            justify=tkinter.CENTER
        )
        # Style.map (TNotebook.Tab)
        style.map(
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
        style.configure(
            "TLabel",
            font=("Meiryo", 30, "bold"),
            foreground="black",
            background="white",
            justify=tkinter.CENTER
        )

        style.configure(
            "TFrame",
            background="white"
        )

        # タブの作成
        tab_list = []
        tab_list.append(tkinter.ttk.Frame(notebook, style="TFrame"))
        tab_list.append(tkinter.ttk.Frame(notebook, style="TFrame"))

        # print(len(notebook))

        # notebookにタブを追加
        notebook.add(tab_list[0], text="本館 2 階")
        notebook.add(tab_list[1], text="本館 4 階")

        # tab_oneに配置するウィジェットの作成
        label = []
        label.append(tkinter.ttk.Label(top_frame, text="座席表", style="TLabel"))
        label.append(tkinter.ttk.Label(
            tab_list[0], text="本館 2 階", style="TLabel"))
        label.append(tkinter.ttk.Label(
            tab_list[1], text="本館 4 階", style="TLabel"))

        # ウィジェットの配置
        top_frame.pack(fill=tkinter.BOTH, padx=10, pady=10)
        notebook.pack(expand=True, fill=tkinter.BOTH,
                      padx=10, pady=10)

        for tmp in label:
            tmp.pack()


def main():
    root = tkinter.Tk()
    app = App(master=root)  # Inherit
    app.mainloop()


main()
