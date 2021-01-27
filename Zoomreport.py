import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time


scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)

client = gspread.authorize(creds)

da = '012121'

# Open spreadhseets
sec1at = client.open("sec1" + da).sheet1
sec2at = client.open("sec2" + da).sheet1
sec1sum = client.open("S21Sec1").sheet1
sec2sum = client.open("S21Sec2").sheet1

sec1att = sec1at.get_all_records()
sec2att = sec2at.get_all_records()
sec1ro = sec1sum.get('A2:A59')
sec2ro = sec2sum.get('A2:A55')

names = sec1at.get('A4:A63')
durations = sec1at.get('B4:B63')
names2 = sec2at.get('A4:A63')
durations2 = sec2at.get('B4:B63')

att1 = dict()
att2 = dict()
sec1ros = list()
sec2ros = list()
cratt1 = list(da)

for r in sec1ro:
    sec1ros.append(r[0])

for n in names:
    for t in durations:
        att1[n[0]] = int(t[0])
        durations.remove(t)
        break

count = 1

row = 2

for i in sec1ros:
    stu = ''

    if i in att1 and att1[i] > 60:
        print(i,"Attended")
        stu = "Attended"
    elif i in att1 and att1[i] > 40:
        print(i,"Late")
        stu = "Late"
    elif i in att1 and att1[i] > 1:
        print(i," Very Late")
        stu = "Very Late"
    else:
        print(i,"No Show")
        stu = "No Show"
    sec1sum.update_cell(row, 4, stu)

    print(count)
    count += 1
    row +=1

#Challenge: google only allow normal users to update 60 times per minute, means if we update one cell at a time, we have to do 2 sessions in two minutes
time.sleep(61)

for r in sec2ro:
    sec2ros.append(r[0])

for n in names2:
    for t in durations2:
        att2[n[0]] = int(t[0])
        durations2.remove(t)
        break

count = 1

row = 2

for i in sec2ros:
    stu = ''
    if i in att2 and att2[i] > 60:
        print(i,att2[i],"Attended")
        stu = "Attended"
    elif i in att2 and att2[i] > 40:
        print(i,att2[i],"Late")
        stu = "Late"
    elif i in att2 and att2[i] > 1:
        print(i,att2[i]," Very Late")
        stu = "Very Late"
    else:
        print(i,"No Show")
        stu = "No Show"
    sec2sum.update_cell(row, 4, stu)
    print(count)
    count += 1
    row +=1