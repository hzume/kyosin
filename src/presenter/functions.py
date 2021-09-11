from datetime import datetime, timedelta

# excel上の日時を変換
def excel_date(num):
    return(datetime(1899, 12, 30) + timedelta(days=num))
