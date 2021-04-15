import openpyxl
import msoffcrypto
import tempfile
from datetime import datetime, timedelta
import PySimpleGUI as sg

# オプションの設定と標準レイアウト
font_name = 'arial'
font_size = 12

frame1 = [
    [
        sg.Text("講師情報", size=(15,1), font=((font_name, font_size))),
    ],
    [   
        sg.Input(size=(20,1), justification="right", font=((font_name, font_size))),
        sg.FileBrowse("ファイルを選択", file_types=(("ALL Files", "*.xlsx"), ), initial_folder="./", font=((font_name, font_size)))
    ],
    [

        sg.Text("管理票", size=(15,1), font=((font_name, font_size))),
    ],
    [   
        sg.Input(size=(20,1), justification="right", font=((font_name, font_size))),
        sg.FileBrowse("ファイルを選択", file_types=(("ALL Files", "*.xlsx"), ), initial_folder="./", font=((font_name, font_size)))
    ],
    [
        sg.Text("テンプレート", size=(15,1), font=((font_name, font_size))),
    ],
    [
        sg.Input(size=(20,1), justification="right", font=((font_name, font_size))),
        sg.FileBrowse("ファイルを選択", file_types=(("ALL Files", "*.xlsx"), ), initial_folder="./", font=((font_name, font_size)))
    ],
    [
        sg.Text("年月", size=(15,1), font=((font_name, font_size)))
    ],
    [
        sg.Input(size=(10,1), justification="right", font=((font_name, font_size))),
        sg.Text("年", font=((font_name, font_size))),
        sg.Input(size=(5,1), justification="right", font=((font_name, font_size))),
        sg.Text("月", font=((font_name, font_size))),
    ],
]

layout = [
   [
       sg.Frame("Input Files", frame1, font=((font_name, font_size)))
   ],
    [
        sg.Button("実行", font=((font_name, font_size)))
    ],
]

# ウィンドウの生成
window = sg.Window('入力ファイルと年月を指定', layout)

# イベントループ
while True:
    event, values = window.read()

    if event is None:
        print('exit')
        break
    
    if event == "実行":
        

# ウィンドウの破棄と終了
window.close()
############################################################################

def excel_date(num):
    return(datetime(1899, 12, 30) + timedelta(days=num))

def main(path_tutor, path_admin, path_template, year, month):
    # 講師情報変更時のみでOK
    # 講師情報を取得
    wb_tutor = openpyxl.load_workbook(path_tutor)
    ws_tutor = wb_tutor.worksheets[0]
    tutors = {}
    tutor_names = []
    for row in ws_tutor.rows:
        if row[0].row == 1:
            keys = row[1:]
        else:
            tutor = {}
            for k, v in zip(keys, row[1:]):
                tutor[k.value] = v.value
            tutor["勤務"] = [0] * 31
            tutors[row[0].value] = tutor
            tutor_names.append(row[0].value)

    ###################################################################################################

    # 管理票を取得
    with open(path_admin, "rb") as f, tempfile.TemporaryFile() as tf:
        ms_file = msoffcrypto.OfficeFile(f)
        ms_file.load_key(password="1219")
        ms_file.decrypt(tf)
        wb = openpyxl.load_workbook(tf, data_only=True)

    # 管理票から対象シートを取得
    trans_table = str.maketrans({"０":"0", "１":"1", "２":"2", "３":"3", "４":"4", "５":"5", "６":"6", "７":"7", "８":"8", "９":"9"})
    sheetnames = []
    for sheetname in wb.sheetnames:
        fixed_name = sheetname.translate(trans_table)
        idx = fixed_name.find("月")
        if idx == -1:
            continue
        if fixed_name[:idx] == str(month):
            sheetnames.append(sheetname)

    # 対象シートから出勤日を取得
    class_times = {"2:00-3:20":0, "3:30-4:50":1, "5:00-6:20":2, "6:30-7:50":3, "8:00-9:20":4}
    for sheetname in sheetnames:
        ws = wb[sheetname]
        MAX_ROW = 100
        MAX_COLUMN = 30
        for i in range(1, MAX_ROW):
            head_cell = ws.cell(row=i, column=2).value
            if (type(head_cell) == int) or (head_cell == None):
                if type(head_cell) == int:
                    date = excel_date(head_cell)
                    # 異なる月の日付があったら警告
                    if date.month != month:
                        break
                    day = date.day - 1 # 0 ~ 30
                table = ws[f"B{i}":f"BA{i+2}"]
                # シート終了の判定
                if table[2][0].value == None: 
                    break
                class_time = class_times[table[2][0].value]
                for j in range(3, MAX_COLUMN):
                    # 登録されていない講師があったらエラー
                    name = table[0][j].value
                    class1 = table[1][j].value
                    class2 = table[2][j].value
                    if name != None and (class1 != None or class2 != None):
                        ##################
                        # 例外コマの処理  #
                        ##################

                        try: 
                            tutors[name]["勤務"][day] |= (1 << class_time)
                        except:
                            print("存在しない講師です")

    ###################################################################################

    # 給与テンプレートを取得
    wb_template = openpyxl.load_workbook(path_template)
    ws_template = wb_template.worksheets[0]
    ws_template["F2"] = year
    ws_template["H2"] = month

    for name in tutor_names:
        tutor = tutors[name]
        fullname = tutor["フルネーム"]
        ws_template["H5"] = fullname
        for day in range(31):
            for class_time in range(5):
                ws_template.cell(10 + day, 4 + class_time).value = 80 if tutor["勤務"][day] & (1 << class_time) else None
        wb_template.save(fullname + f"{year}年{month}月.xlsx")
    