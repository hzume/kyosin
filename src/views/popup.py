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
                password = int(values["password"])
                break
        self.window.close()
        return password