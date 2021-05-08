import PySimpleGUI as sg
import openpyxl
import msoffcrypto
import tempfile
from src.views.view import InterFace
from src.models.model import Tutor, Tutors, Class
from src.settings.translate import *

class Presenter:
    def __init__(self, window: sg.Window):
        self.window = window
        self.tutors = Tutors()
    
    # ---to Model---
    def tutors_init(self, params_list):
        for params in params_list:
            self.tutors.add(Tutor(**params))
    
    def 

    # ---from View---
    def get_tutors(self, tutor_path):
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
    
