from typing import ValuesView
import PySimpleGUI as sg
import warnings
from src.presenter.presenter import Presenter
from src.presenter.handler import Handler
from src.views.view import InterFace

#warnings.simplefilter("ignore")

def main():
    interface = InterFace()
    presenter = Presenter(window=interface.window)
    handler = Handler(presenter)

    while True:
        event, values = interface.window.read()
        handler.handle(event, values)

        if event is None:
            break

    interface.close_window()
    
if __name__=="__main__":
    main()