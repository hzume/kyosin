class Tutor:
    def __init__(self, name, fullname, pay_class, pay_officework, trans_fee):
        self.name = name
        self.fullname = fullname
        self.pay_class = pay_class
        self.pay_officework = pay_officework
        self.trans_fee = trans_fee
        self.workday = [0] * 31
    
    def workday_update(self, day, class_time):
        self.workday[day] |= (1 << class_time)
    
class Class:
    def __init__(self, day, class_time, name, officework=False):
        self.day = day
        self.class_time = class_time
        self.name = name
        self.officework = officework
        
