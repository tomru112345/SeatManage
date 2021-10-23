# 自習室管理アプリケーション

## EXE 化コマンドについて

```bash
pyinstaller ./ListBox_tk.py --onefile --icon="../image/python_LOGO.ico" --noconsole --name="SeatManage" --hidden-import="openpyxl,pkg_resources.py2_warn,importlib"
```