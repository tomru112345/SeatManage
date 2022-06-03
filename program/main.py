#!python3.8
import Seat
import settings
import tkinter


def webView():
    """TODO: webview を使用して GUI を作る"""
    import webview
    webview.create_window('Hello world', 'https://google.com')
    webview.start(gui='cef')

def main():
    root = tkinter.Tk()
    # root.overrideredirect(1)
    root.title('座席表')
    root.geometry("1920x1080")
    # root.state("zoomed")
    root.iconbitmap(settings.pythonLOGOICO)
    Seat.Seat(root)
    root.mainloop()

if __name__ == '__main__':
    main()


