import json
import os
import openpyxl


class Reader():
    # 初期変数
    def __init__(self):
        # ファイルの配置場所
        self.JSON_PATH = "./settings.json"
        self.EXCEL_PATH = "../生徒名簿.xlsx"

        # エンコード形式
        self.EncodeItem = 'utf-8_sig'

        self.data = None

        self.HEIGHT = 7
        self.LENGTH = 5
        self.IgnoreLists = None
        self.DicYear = {}

        # 学生リスト
        self.student_list = []

        # コースのリスト
        self.COURSE = []

        # ソートされた結果のリスト
        self.choose_list = []

        self.yearName = []
        self.School_year_ID = []
        self.student_list_keys = []
        self.LAYOUT = []

        self.JsonRead()
        self.ExcelRead()
        pass

    def JsonRead(self) -> None:
        if os.path.isfile(self.JSON_PATH):
            with open(self.JSON_PATH, 'r', encoding=self.EncodeItem) as f:
                self.data = json.load(f)
                self.EXCEL_PATH = self.data['FilePath']['StudentsListPath']
                self.HEIGHT = int(self.data['SeatSettings']['height'])
                self.LENGTH = int(self.data['SeatSettings']['length'])
                self.IgnoreLists = self.data['SeatSettings']['IgnoreSeat']
                for i in self.data['Year']:
                    self.DicYear[i] = self.data['Year'][i]

                # 2次元配列のとおりに、gridでレイアウトを作成する
                for y in range(self.HEIGHT):
                    tmp_mini_list = []
                    for x in range(self.LENGTH):
                        tmp_mini_list.append('-')
                    self.LAYOUT.append(tmp_mini_list)
                for tmp_btn in self.IgnoreLists:
                    if (0 <= tmp_btn['height'] and tmp_btn['height'] < self.HEIGHT) and (0 <= tmp_btn['length'] and tmp_btn['length'] < self.LENGTH):
                        self.LAYOUT[tmp_btn['height']][tmp_btn['length']] = 'x'
        pass

    def JsonWrite(self) -> None:
        if os.path.isfile(self.JSON_PATH):
            # 更新内容
            self.data['FilePath']['StudentsListPath'] = self.EXCEL_PATH
            self.data['SeatSettings']['height'] = self.HEIGHT
            self.data['SeatSettings']['length'] = self.LENGTH

            for i in self.DicYear.keys():
                self.data['Year'][i] = self.DicYear[i]

            tmp_ignore_list = []
            if (len(self.IgnoreLists) > 0):
                for tmp_item in self.IgnoreLists:
                    tmp_ignore_dict = {
                        "height": tmp_item['height'], "length": tmp_item['length']}
                    tmp_ignore_list.append(tmp_ignore_dict)
            self.data['SeatSettings']['IgnoreSeat'] = tmp_ignore_list

            # Pythonオブジェクトをファイル書き込み
            with open(self.JSON_PATH, 'w', encoding=self.EncodeItem) as f:
                json.dump(self.data, f, indent=4, ensure_ascii=False)
        pass

    def ExcelRead(self) -> bool:
        """Excel ファイルのロード"""
        if os.path.isfile(self.EXCEL_PATH):
            wb = openpyxl.load_workbook(self.EXCEL_PATH)

            self.yearName = []
            self.School_year_ID = []

            for ws in wb.worksheets:
                year = ws.title
                self.yearName.append(year)
                for row in ws.rows:
                    if row[0].row == 1:
                        # １行目
                        check_header = []
                        header_cells = row
                        for cell in row:
                            check_header.append(cell.value)
                        if ['在籍番号', '氏名', 'カナ名', 'コース'] != check_header:
                            return False
                    else:
                        # ２行目以降
                        row_dic = {}
                        # セルの値を「key-value」で登録
                        for k, v in zip(header_cells, row):  # zip 関数(forループで複数のリストの要素を取得)
                            if not (k.value == '電話番号' or k.value == '在籍' or k.value == '面談'):
                                row_dic[k.value] = v.value
                        self.student_list.append(row_dic)

            self.student_list_keys = list(
                self.student_list[0].keys())  # Excel の表のキー取得
            for i in self.student_list:
                t = str(i[self.student_list_keys[0]])
                self.School_year_ID.append(t[0:2])
                self.COURSE.append(str(i[self.student_list_keys[3]]))
            self.School_year_ID = sorted(list(set(self.School_year_ID)))
            self.COURSE = sorted(list(set(self.COURSE)))
            self.setDicYear(self.yearName, self.School_year_ID)
            self.JsonWrite()
            return True
        else:
            return False

    def setDicYear(self, yearName, School_year_ID) -> None:
        newDicYear = {}
        for i in range(len(yearName)):
            newDicYear[yearName[i]] = School_year_ID[i]
        self.DicYear = newDicYear
        pass

    def choose_CYname(self, choose_COURSE, choose_year, kana_num):  # 選択するコース,学年のリストの要素数
        """生徒の選択"""
        kana_list = [
            ["ア", "イ", "ウ", "エ", "オ", "ｱ", "ｲ", "ｳ", "ｴ", "ｵ"],
            ["カ", "キ", "ク", "ケ", "コ", "ｶ", "ｷ", "ｸ", "ｹ", "ｺ",
             "ガ", "ギ", "グ", "ゲ", "ゴ", "ｶﾞ", "ｷﾞ", "ｸﾞ", "ｹﾞ", "ｺﾞ"],
            ["サ", "シ", "ス", "セ", "ソ", "ｻ", "ｼ", "ｽ", "ｾ", "ｿ",
             "ザ", "ジ", "ズ", "ゼ", "ゾ", "ｻﾞ", "ｼﾞ", "ｽﾞ", "ｾﾞ", "ｿﾞ"],
            ["タ", "チ", "ツ", "テ", "ト", "ﾀ", "ﾁ", "ﾂ", "ﾃ", "ﾄ",
             "ダ", "ヂ", "ヅ", "デ", "ド", "ﾀﾞ", "ﾁﾞ", "ﾂﾞ", "ﾃﾞ", "ﾄﾞ"],
            ["ナ", "ニ", "ヌ", "ネ", "ノ", "ﾅ", "ﾆ", "ﾇ", "ﾈ", "ﾉ"],
            ["ハ", "ヒ", "フ", "ヘ", "ホ", "ﾊ", "ﾋ", "ﾌ", "ﾍ", "ﾎ",
             "バ", "ビ", "ブ", "ベ", "ボ", "ﾊﾞ", "ﾋﾞ", "ﾌﾞ", "ﾍﾞ", "ﾎﾞ",
             "パ", "ピ", "プ", "ペ", "ポ", "ﾊﾟ", "ﾋﾟ", "ﾌﾟ", "ﾍﾟ", "ﾎﾟ"],
            ["マ", "ミ", "ム", "メ", "モ", "ﾏ", "ﾐ", "ﾑ", "ﾒ", "ﾓ"],
            ["ヤ", "ユ", "ヨ", "ﾔ", "ﾕ", "ﾖ"],
            ["ラ", "リ", "ル", "レ", "ロ", "ﾗ", "ﾘ", "ﾙ", "ﾚ", "ﾛ"],
            ["ワ", "ヲ", "ン", "ﾜ", "ｦ", "ﾝ"]
        ]
        choose_list = []
        kana_list = kana_list[kana_num]
        for i in self.student_list:
            if i[self.student_list_keys[3]] == choose_COURSE:
                if str(i[self.student_list_keys[0]]).startswith(self.DicYear[choose_year]):
                    if (i[self.student_list_keys[2]])[0:1] in kana_list:
                        choose_list.append(i[self.student_list_keys[1]])
        choose_list = list(set(choose_list))
        choose_list = sorted(choose_list)
        return choose_list
