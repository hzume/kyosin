import PySimpleGUI as sg
from src.views.style import *
from src.settings.config import *

class InterFace:
    def __init__(self):
        self.setting_frame = [sg.Frame("Setting", **text_style, layout=[
            [
                sg.Text("講師情報", **text_style, **input_text_style),
            ],
            [   
                sg.Input(default_text=default.get("tutor_path"), key="tutor_path", **input_form_style, **text_style),
                sg.FileBrowse("参照", file_types=(("ALL Files", "*.xlsx"), ), initial_folder="./", **text_style)
            ],
            [

                sg.Text("管理票", **text_style, **input_text_style),
            ],
            [   
                sg.Input(default_text=default.get("admin_path"), key="admin_path", **input_form_style, **text_style),
                sg.FileBrowse("参照", file_types=(("ALL Files", "*.xlsx"), ), initial_folder="./", **text_style)
            ],
            [
                sg.Text("テンプレート", **text_style, **input_text_style),
            ],
            [
                sg.Input(default_text=default.get("template_path"), key="template_path", **input_form_style, **text_style),
                sg.FileBrowse("参照", file_types=(("ALL Files", "*.xlsx"), ), initial_folder="./", **text_style)
            ],
            [
                sg.Text("年月", **text_style, **input_text_style)
            ],
            [
                sg.Input(default_text=default.get("year"), size=(10,1), key="year", enable_events=True, justification="right", **text_style),
                sg.Text("年", **text_style),
                sg.Input(size=(5,1), key="month", enable_events=True, justification="right", **text_style),
                sg.Text("月", **text_style),
            ],
            [
                sg.Text("出力先", **text_style, **input_text_style)
            ],
            [
                sg.Input(default_text=default.get("output_folder"), key="output_folder", **input_form_style, **text_style),
                sg.FolderBrowse("参照", initial_folder="../", **text_style)
            ]
        ])]

        self.log_frame = [sg.Frame("Log", **text_style, layout=[
            [
                sg.Output(size=(50, 17), **text_style)
            ]
        ])]

        self.layout = [
            self.input_frame,
            self.output_frame,
            sg.Button("実行", **text_style),
            sg.Button("ミーティング設定", **text_style)
        ]

    def show_window(self):
        sg.theme('SystemDefault')
        self.window = sg.Window("給与明細作成", layout=self.layout)   

    def close_window(self):
        self.window.close()
    