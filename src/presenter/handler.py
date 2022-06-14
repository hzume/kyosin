from src.presenter.presenter import Presenter

class Handler:
    def __init__(self, presenter: Presenter):
        self.presenter = presenter
        self.func_dict = {
            "exec": self.presenter.exec,
            "add_meeting": self.presenter.meeting_setting,
            "list_meeting": self.presenter.list_meeting,
            "del_meeting": self.presenter.del_meeting
        }
    
    def handle(self, event_key, values):
        if event_key not in self.func_dict:
            return
        try:
            event_func = self.func_dict[event_key]
            event_func(values)
        except Exception:
            import traceback
            traceback.print_exc()
            return
