import PySimpleGUI as sg
from src.views.style import *
from src.settings.config import *

class InterFace:
    def __init__(self):
        sg.theme('SystemDefault')
        
        self.setting_frame = sg.Frame("Setting", **text_style, layout=[
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
                sg.Input(default_text=default.get("year"), key="year", **short_form_style, **text_style),
                sg.Text("年", **text_style),
                sg.Input(key="month", **short_form_style, **text_style),
                sg.Text("月", **text_style),
            ],
            [
                sg.Text("出力先", **text_style, **input_text_style)
            ],
            [
                sg.Input(default_text=default.get("output_folder"), key="output_folder", **input_form_style, **text_style),
                sg.FolderBrowse("参照", initial_folder="../", **text_style)
            ]
        ])

        self.log_frame = sg.Frame("Log", **text_style, layout=[
            [
                sg.Output(size=(50, 17), **text_style)
            ]
        ])

        self.layout = [
            [
                self.setting_frame,
                self.log_frame,
            ],
            [    
                sg.Button("実行", key="exec", **text_style),
                sg.Button("ミーティング設定", key="set_meeting", **text_style)
            ]
        ]
        self.window = sg.Window("給与明細作成", layout=self.layout)

    def close_window(self):
        self.window.close()

