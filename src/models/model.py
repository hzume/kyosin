from src.settings.translate import *

class Class:
    def __init__(self, day, class_time, name, officework=False, officework_time=80):
        self.day = day
        self.class_time = class_time
        self.name = name
        self.officework = officework

class Tutor:
    def __init__(self, name, fullname, pay_class, pay_officework, trans_fee, type):
        self.name = name
        self.fullname = fullname
        self.pay_class = pay_class
        self.pay_officework = pay_officework
        self.trans_fee = trans_fee
        self.class_work = [0] * 31
        self.office_work = [0] * 31
        self.meeting = [0] * 31
        self.type = type
    
    def worktime_update(self, clas: Class):
        if clas.officework:
            self.office_work[clas.day] += class_length
        else:
            self.class_work[clas.day] |= (1 << clas.class_time)
    
class Tutors(dict):
    def add(self, tutor: Tutor):
        self[tutor.name] = tutor

    def worktime_update(self, clas: Class):
        try:
            self[clas.name].worktime_update(clas)
        except:
            print("存在しない講師です: " + clas.name)

class Meeting:
    def __init__(self, meeting_day, meeting_length, participants):
        self.meeting_day = meeting_day
        self.meeting_length = meeting_length
        self.participants = participants
