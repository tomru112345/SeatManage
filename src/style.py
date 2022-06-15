import tkinter
import tkinter.ttk


def set_style(style: tkinter.ttk.Style) -> None:
    # スタイルの設定
    style.theme_use('clam')

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
        "header.TLabel",
        font=("Meiryo", 30, "bold"),
        foreground="black",
        background="white",
        justify=tkinter.CENTER
    )

    style.configure(
        "h1.TLabel",
        font=("Meiryo", 30, "bold"),
        foreground="white",
        background="SteelBlue2",
        side=tkinter.TOP
    )

    style.configure(
        "h2.TLabel",
        font=("Meiryo", 24, "bold"),
        foreground="white",
        background="indianred1",
        side=tkinter.TOP
    )

    style.configure(
        "h3.TLabel",
        font=("Meiryo", 18, "bold"),
        foreground="white",
        background="MediumPurple1",
        side=tkinter.TOP
    )

    style.configure(
        "p.TLabel",
        font=("Meiryo", 15),
        foreground="black",
        background="white",
        side=tkinter.TOP
    )

    style.configure(
        "p.bold.TLabel",
        font=("Meiryo", 15, "bold"),
        foreground="red",
        background="white",
        side=tkinter.TOP
    )

    style.configure(
        "error.h1.TLabel",
        font=("Meiryo", 30, "bold"),
        foreground="red",
        background="white",
        justify=tkinter.CENTER
    )

    style.configure(
        "TFrame",
        background="white"
    )

    # Style
    style.configure(
        "TButton",
        font=("Meiryo", 30, "bold"),
        # width=10,
        background="medium spring green",
        foreground="black"
        # justify=tkinter.CENTER
    )

    style.map(
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
