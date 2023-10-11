from src.settings.translate import *

class Class:
    def __init__(self, day:int, class_time:int, name:str, officework=0, officework_time=80) -> None:
        self.day = day
        self.class_time = class_time
        self.name = name
        self.officework = officework
        self.officework_time = officework_time

class Tutor:
    def __init__(self, name: str, fullname: str, pay_class: float, pay_officework: float, trans_fee: float, type: str) -> None:
        self.name = name
        self.fullname = fullname
        self.pay_class = pay_class
        self.pay_officework = pay_officework
        self.trans_fee = trans_fee
        self.class_work = [0] * 31
        self.office_work = [0] * 31
        self.meeting = [0] * 31
        self.type = type
    
    def worktime_update(self, _class: Class) -> None:
        if _class.officework:
            self.office_work[_class.day] += _class.officework_time
        else:
            self.class_work[_class.day] |= (1 << _class.class_time)
    
class Tutors(dict[str, Tutor]):
    def add(self, tutor: Tutor) -> None:
        self[tutor.name] = tutor

    def worktime_update(self, _class: Class) -> None:
        try:
            self[_class.name].worktime_update(_class)
        except:
            print("存在しない講師です: " + _class.name)

class Meeting:
    def __init__(self, meeting_day:int, meeting_length:int, participants:list[Tutor]):
        self.meeting_day = meeting_day
        self.meeting_length = meeting_length
        self.participants = participants
