import PySimpleGUI as sg
from src.views.view import InterFace
from src.models.model import Tutor, Class

class Presenter:
    def __init__(self, window: sg.Window):
        self.window = window
        self.tutor = {}

    # --------to Model--------
    def new_tutor(self, name, fullname, pay_class, pay_officework, trans_fee):
        self.tutor[name] = Tutor(name, fullname, pay_class, pay_officework, trans_fee)
    
    # --------from Model--------
    
    # --------to View--------
    