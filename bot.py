import requests
from lxml import html
import xlrd
# import xlwt
# from xlutils.copy import copy



url = "http://info.vit.ac.in/gravitas18/gravitas/gravitas_coordinator_login.asp"
session = requests.Session()
session.cookies.clear()
result = session.get(url)
tree = html.fromstring(result.text)

captcha = str(html.tostring(tree.xpath("/html/body/div/div[3]/div[2]/div/div[2]/form/div[3]/div[1]/font")[0]))
captcha = captcha[captcha.find("</font>")-6:captcha.find("</font>")]
cookie = session.cookies.keys()[0] + "=" + session.cookies.values()[0]

headers = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36',
    'cookie': cookie
}
payload = {
    'loginid': 'YOUR_EMAILD_ID_HERE',
    'logpassword': 'YOUR_PASSWORD_HERE',
    'captchacode1': captcha,
    'captchacode': captcha,
    'frmSubmit': ''
}
post_url = "http://info.vit.ac.in/gravitas18/gravitas/coord_login_authorize.asp"
req = session.post(post_url, data=payload, headers=headers)

if req.status_code == 200:
    print("Successfully logged in!")
else:
    quit(0)

sheet = xlrd.open_workbook('attendance.xlsx').sheet_by_index(0)
total_attendance_status = [0, 0, 0, 0, 0]
absentees = []
presentees = []

for i in range(9, sheet.nrows):
    count = 0
    for j in range(4, 8):
        if str(sheet.cell_value(i, j)).lower() == 'p':
            count += 1
    total_attendance_status[count] = total_attendance_status[count] + 1
    if count != 4:
        absentees.append(str(sheet.cell_value(i, 1)))
    else:
        presentees.append(str(sheet.cell_value(i, 1)))

# Update the event ID with your event ID

for present in presentees:
    post_url = "http://info.vit.ac.in/gravitas18/gravitas/coord_event_post_attendance.asp"
    payload = {
        'stdid': present,
        'eveid': 'YOUR_EVENT_ID_HERE',
        'upstatus': 'true',
        'attdStatus': 'Present',
        'winStatus': '',
        'frmSubmit55': ''
    }
    req = session.post(post_url, data=payload, headers=headers)
    print("Number: {}\tCode: {}".format(present, req.status_code))

print("Posting for absentees...")
for absent in absentees:
    post_url = "http://info.vit.ac.in/gravitas18/gravitas/coord_event_post_attendance.asp"
    payload = {
        'stdid': absent,
        'eveid': 'YOUR_EVENT_ID_HERE',
        'upstatus': 'true',
        'attdStatus': 'Absent',
        'winStatus': '',
        'frmSubmit55': ''
    }
    req = session.post(post_url, data=payload, headers=headers)
    print("Number: {}\tCode: {}".format(absent, req.status_code))

for i in range(0, 5):
    print("Number of people present for {} sessions are {}".format(i, total_attendance_status[i]))

