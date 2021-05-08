import PySimpleGUI as sg
import openpyxl
import msoffcrypto
import tempfile
from os.path import basename
from datetime import datetime, timedelta
from src.views.view import InterFace
from src.models.model import Tutor, Tutors, Class
from src.settings.translate import *
from src.settings.config import *

class PasswordMistake(Exception):
    pass

class Presenter:
    def __init__(self, window: sg.Window):
        self.window = window
        self.tutors = Tutors()

    def excel_date(num):
        return(datetime(1899, 12, 30) + timedelta(days=num))
    
    # ---to Model---
    def init_tutors(self, params_list):
        for params in params_list:
            self.tutors.add(Tutor(**params))

    def set_worktime(self, params_list):
        for params in params_list:
            self.tutors.worktime_update(Class(**params))

    def update_config(self, values):
        default["tutor_path"] = values["tutor_path"]
        default["admin_path"] = values["admin_path"]
        default["template_path"] = values["template_path"]
        default["year"] = values["year"]
        default["output_folder"] = values["output_folder"]

    # ---from View---
    def receive_tutors(self, tutor_path):
        wb = openpyxl.load_workbook(tutor_path)
        ws = wb.worksheets[0]
        params_list = []
        for row in ws.rows:
            if row[0].row == 1:
                keys = row
            else:
                params = {}
                for k, v in zip(keys, row):
                    params[ja_to_en[k.value]] = v.value
                params_list.append(params)
        return params_list
    
    def receive_classes(self, admin_path, month, password, MAX_ROW = 100, MAX_COLUMN = 30):
        if password != default[password]:
            raise Exception
        params_list = []
        with open(admin_path, "rb") as f, tempfile.TemporaryFile() as tf:
            ms_file = msoffcrypto.OfficeFile(f)
            ms_file.load_key(password=password)
            ms_file.decrypt(tf)
            wb = openpyxl.load_workbook(tf, data_only=True)
        trans_table = str.maketrans(zenkaku_to_hankaku)
        sheetnames = []

        for sheetname in wb.sheetnames:
            fixed_name = sheetname.translate(trans_table)
            idx = fixed_name.find("月")
            if idx == -1:
                continue
            if fixed_name[:idx] == str(month):
                sheetnames.append(sheetname)

        for sheetname in sheetnames:
            ws = wb[sheetname]
            for i in range(1, MAX_ROW):
                head_cell = ws.cell(row=i, column=2).value
                if (type(head_cell) == int) or (head_cell == None):
                    if type(head_cell) == int:
                        date = self.excel_date(head_cell)
########################## 注意 #########################
                        if date.month != month:
                            break                     
                        day = date.day - 1 # 0 ~ 30
                    table = ws[f"B{i}":f"BA{i+2}"]
                    # シート終了の判定
                    if table[2][0].value == None: 
                        break
                    class_time = class_times[table[2][0].value]
                    for j in range(3, MAX_COLUMN):
                        params = {}
                        name = table[0][j].value
                        class1 = table[1][j].value
                        class2 = table[2][j].value
                        if name != None and (class1 != None or class2 != None):
                            params["day"] = day
                            params["name"] = name
                            params["class_time"] = class_time
                            if class1 == "事務" or class2 == "事務":
                                params["office_work"] = True
                            else:
                                params["office_work"] = False
                        params_list.append(params)
    
    # ---to View---
    def make_payslip(self, template_path, year, month, output_folder):
        wb = openpyxl.load_workbook(template_path)
        ws = wb.worksheets[0]
        ws["F2"] = year
        ws["H2"] = month

        for tutor in self.tutors.values():
            office_time = 0
            if tutor.name == None:
                continue
            ws["H5"].value = tutor.fullname
            ws["K41"].value = f"=ROUNDDOWN(SUM(K10:K40)/60*{tutor.pay_class},0)"
            ws["L41"].value = f"=ROUNDDOWN(SUM(L10:L40)/60*{tutor.pay_office},0)"
            ws["R41"].value = f'=COUNTIF(R10:R40,"○")*{tutor.trans_fee}'
            for day in range(31):
                for class_time in range(5):
                    if tutor.class_work[day] & (1 << class_time):
                        ws.cell(row=10+day, column=4+class_time).value = 80
                        office_time += 30
                    if tutor.office_work[day] & (1 << class_time):
                        office_time += 80
                ws[f"L{10+day}"].value = office_time
            wb.save(output_folder + "/" + tutor.fullname + f"{year}年{month}月.xlsx")
            print(tutor.fullname + f"{year}年{month}月.xlsx を出力")
                
    # ---Event Process---
    def exec(self, values):
        if None in values.values():
            print("入力されていない項目があります")
            return
        self.update_config(values)
        print("処理を実行")
        print("--処理対象ファイル--")
        print(" 講師情報 　　:" + basename(values["tutor_path"]))
        print(" 管理票 　　　:" + basename(values["admin_path"]))
        print(" テンプレート :" + basename(values["template_path"]))
        print("------------------")
        tutors_params = self.receive_tutors(**values)
        self.init_tutors(tutors_params)
        try:
            classes_params = self.receive_classes(**values)
        except PasswordMistake:
            print("Error:パスワードが間違っています")
        except Exception as e:
            print("Error:不明なエラーです")
            print(e)
        self.set_worktime(classes_params)
        self.make_payslip(**values)
        print("処理を終了しました")

