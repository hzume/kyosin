import PySimpleGUI as sg
from src.views.style import *
from src.settings.config import *

class PasswordPopup:
    def __init__(self):
        sg.theme('SystemDefault')
        self.layout = [
            [
                sg.Text("パスワードを入力", **input_text_style, **text_style)
            ],
            [
                sg.Input(key="password", password_char="*", **short_form_style, **text_style), 
                sg.Button("送信", key="input_password", **text_style)
            ]
        ]
        self.window = sg.Window("パスワード", layout=self.layout, auto_size_text=True, auto_size_buttons=True)
    
    def receive_password(self):
        while True:
            event, values = self.window.read()
            if event == None:
                password = None
                break
            if event == "input_password":
                password = values["password"]
                break
        return password
    
    def close_window(self):
        self.window.close()

class MeetingPopup:
    def __init__(self, tutor_names: list):
        sg.theme('SystemDefault')
        self.layout = [
            [
                sg.Text("実施日", **input_text_style, **text_style)
            ],
            [
                sg.Input(key="meeting_day", **short_form_style, **text_style), 
                sg.Text("日", **text_style)
            ],
            [
                sg.Text("時間", **input_text_style, **text_style)
            ],
            [
                sg.Input(key="meeting_length", **short_form_style, **text_style),
                sg.Text("分", **text_style)
            ]
        ] + \
        [
            [
                sg.Checkbox(text=name, key=name, **text_style)
            ] for name in tutor_names
        ] + \
        [
            [
                sg.OK()
            ]
        ]
        self.window = sg.Window("ミーティング設定", layout=self.layout)
        self.tutor_names = tutor_names
        self.meeting_day = ""
        self.meeting_length = ""
        self.participants = []
    
    def receive_meeting(self): 
        while True:
            event, values = self.window.read()
            if event == None:
                break

            elif event == "OK":
                for name in self.tutor_names:
                    if values[name]:
                        self.participants.append(name)
                self.meeting_day = values["meeting_day"]
                self.meeting_length = values["meeting_length"]
                break
                
        return self.meeting_day, self.meeting_length, self.participants

    def close_window(self):
        self.window.close()
