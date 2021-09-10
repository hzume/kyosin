import re

s1 = "事務60"
s2 = ""
time1 = re.findall(r"[0-9]+", s1)
time2 = re.findall(r"[0-9]+", s2)
time1 = int(time1[0]) if time1 else 0
time2 = int(time2[0]) if time2 else 0
print(time1, time2)