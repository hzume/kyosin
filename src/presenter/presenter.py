import PySimpleGUI as sg
import openpyxl
import msoffcrypto
import tempfile
from os.path import basename
from datetime import datetime, timedelta
from src.views.view import InterFace
from src.views.popup import PasswordPopup, MeetingPopup
from src.models.model import Tutor, Tutors, Class
from src.settings.translate import *
from src.settings.config import *

class Presenter:
    def __init__(self, window: sg.Window):
        self.window = window
        self.tutors = Tutors()

    def excel_date(self, num):
        return(datetime(1899, 12, 30) + timedelta(days=num))

    # ---from Model---
    def get_tutor_names(self):
        tutor_names = []
        for name in self.tutors.keys():
            if name != None:
                tutor_names.append(name)
        return tutor_names
    
    # ---to Model---
    def init_tutors(self, params_list):
        for params in params_list:
            self.tutors.add(Tutor(**params))

    def set_worktime(self, params_list):
        for params in params_list:
            self.tutors.worktime_update(Class(**params))

    def update_config(self, values):
        with open("src/settings/config.py", "w") as f:
            f.write(
f"""default = {{
    'password': 1219,
    'tutor_path': '{values["tutor_path"]}',
    'admin_path': '{values["admin_path"]}',
    'template_path': '{values["template_path"]}',
    'year': {values["year"]},
    'output_folder': '{values["output_folder"]}',
}}
"""
            )
    
    def set_meeting(self, meeting_day, meeting_length, participants):
        for participant in participants:
            self.tutors[participant].office_work[meeting_day] += meeting_length

    # ---from View---
    def receive_meeting(self, tutor_names):
        popup = MeetingPopup(tutor_names)
        meeting_day, meeting_length, participants = popup.receive_meeting()
        if meeting_day == "" or meeting_length == "":
            raise Exception
        popup.close_window()
        del popup
        return meeting_day, meeting_length, participants

    def receive_password(self):
        popup = PasswordPopup()
        password = popup.receive_password()
        if password == "":
            raise Exception
        popup.close_window()
        del popup
        return int(password)
        
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
        params_list = []
        with open(admin_path, "rb") as f, tempfile.TemporaryFile() as tf:
            ms_file = msoffcrypto.OfficeFile(f)
            ms_file.load_key(password=str(password))
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
                                params["officework"] = True
                            else:
                                params["officework"] = False
                            params_list.append(params)
        return params_list
    
    # ---to View---
    def make_payslip(self, template_path, year, month, output_folder):
        wb = openpyxl.load_workbook(template_path)
        ws = wb.worksheets[0]
        ws["F2"] = year
        ws["H2"] = month

        for tutor in self.tutors.values():
            if tutor.name == None:
                continue
            ws["H5"].value = tutor.fullname
            ws["K41"].value = f"=ROUNDDOWN(SUM(K10:K40)/60*{tutor.pay_class},0)"
            ws["L41"].value = f"=ROUNDDOWN(SUM(L10:L40)/60*{tutor.pay_officework},0)"
            ws["R41"].value = f'=COUNTIF(R10:R40,"○")*{tutor.trans_fee}'
            for day in range(31):
                for class_time in range(5):
                    if tutor.class_work[day] & (1 << class_time):
                        ws.cell(row=10+day, column=4+class_time).value = class_length
                        tutor.office_work[day] += officetime_per_class 
                ws[f"L{10+day}"].value = tutor.office_work[day]
            wb.save(output_folder + "/" + tutor.fullname + f"{year}年{month}月.xlsx")
            print(tutor.fullname + f"{year}年{month}月.xlsx を出力")
                
    # ---Event Process---
    def exec(self, values):
        if "" in [values["tutor_path"], values["admin_path"], values["template_path"], values["year"], values["month"], values["output_folder"]]:
            print("入力されていない項目があります")
            return
        self.update_config(values)
        try:
            password = self.receive_password()
        except:
            print("パスワードが入力されていません")
            return
        if password != default["password"]:
            print("パスワードが異なります")
            return
        print("処理を実行")
        print("----処理対象ファイル--------------------")
        print("| 講師情報 　　|" + basename(values["tutor_path"]))
        print("| 管理票 　　　|" + basename(values["admin_path"]))
        print("| テンプレート |" + basename(values["template_path"]))
        print("---------------------------------------")
        tutors_params = self.receive_tutors(values["tutor_path"])
        self.init_tutors(tutors_params)
        try:
            classes_params = self.receive_classes(values["admin_path"], int(values["month"]), password)
        except:
            import traceback
            traceback.print_exc()
            return
        self.set_worktime(classes_params)
        self.make_payslip(values["template_path"], int(values["year"]), int(values["month"]), values["output_folder"])
        print("処理を終了しました")

    def meeting_setting(self, values):
        tutor_path = values["tutor_path"]
        if tutor_path == None:
            print("講師情報が指定されていません")
            return
        tutors_params = self.receive_tutors(tutor_path)
        self.init_tutors(tutors_params)
        tutor_names = self.get_tutor_names()
        try:
            meeting_day, meeting_length, participants = self.receive_meeting(tutor_names)
        except:
            print("入力されていない項目があります")
            return
        self.set_meeting(int(meeting_day), int(meeting_length), participants)
        print("a")
