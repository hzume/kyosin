import PySimpleGUI as sg
from src.views.style import *
from src.views.error import *
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
                raise PassError
            if event == "input_password":
                password = values["password"]
                break
        self.window.close()
        return password

class MeetingPopup:
    def __init__(self, tutor_names: list):
        sg.theme('SystemDefault')
        cols = [[sg.Checkbox(text=name, key=name, **text_style)] for name in tutor_names]

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
            ],
            [
                sg.Column(cols, scrollable=True, vertical_scroll_only=True, size=(200,400))
            ],
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
                raise PassError

            elif event == "OK":
                for name in self.tutor_names:
                    if values[name]:
                        self.participants.append(name)
                self.meeting_day = values["meeting_day"]
                self.meeting_length = values["meeting_length"]
                break
        self.window.close()       
        return self.meeting_day, self.meeting_length, self.participants
        
