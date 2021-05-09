from src.presenter.presenter import Presenter

class Handler:
    def __init__(self, presenter: Presenter):
        self.presenter = presenter
        self.func_dict = {
            "exec": self.presenter.exec,
            "set_meeting": self.presenter.meeting_setting
        }
    
    def handle(self, event_key, values):
        if event_key not in self.func_dict:
            return
        event_func = self.func_dict[event_key]
        event_func(values)
