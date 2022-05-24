# -*- coding: UTF-8 -*-

import xlrd
import json
import random
import requests
import win32com.client as win32

data = {
  'client_id': 'LrfBSmE61wqUGpxlQhK6IK9YbuCwjmH8ZLzRVT1lXUU',
  'client_secret': 'c2W2ug8UHP28B2rcFDJrOlBHVpHx3e4rPpwFuhHVE2U',
  'grant_type': "client_credentials"
}

r = requests.post("https://www.showmebug.com/oauth/token.json", data=data)

headers = {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer {}'.format(r.json().get('access_token'))
}


EXAM_PADS = {
  'EE Online Test B': '110346',
  'EE Online Test C': '110795'
}

exam_id = random.choice(list(EXAM_PADS.values()))
print('Selecting {}'.format(exam_id))

data = {
  'exam_id': exam_id,
  'candidates': []
}

candidates = {}

wb = xlrd.open_workbook('SW_template.xlsx')
sh = wb.sheet_by_name('SW')
for i in range(sh.nrows):
    cur_row = sh.row_values(i)
    if i != 0 and len(cur_row[1]) > 0:
        name = sh.row_values(i)[0].strip()
        data.get('candidates').append({'uid': ''.join(random.sample('0123456789', 5)), 'name': name})
        candidates[name] = sh.row_values(i)[1]


data = json.dumps(data)

r = requests.post("https://www.showmebug.com/open_api/v1/batch_written_pads", headers=headers, data=data)
print(r.text)
response = r.json()

if response.get('errcode') == 0:
    for pad in response.get('data').get('written_pads'):
        print(pad)
        name = pad['candidate_name']
        print('Sending email to {}'.format(name))
        email = candidates[name]
        interviewer_url = pad['url']

        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = email
        mail.Subject = 'Siemens Efficient Engineering SDE Opportunity - ACTION REQUIRED'
        mail.HTMLBody = '<p>Hello {},</p><p>Thank you for your interest in Siemens&#39; Efficient Engineering SDE opportunities! We would like to invite you to complete our Online Assessment, the first step of our interview process. To be considered for interviews, please complete the assessment <strong>no later than 2 days from now, by 11:59PM CST</strong>. </p><p><strong>Online Assessment Overview</strong></p><p>The assessment consists of two components: two short essay questions (~10 mins each), and a coding problem (~100 mins). Please spend your time wisely in case of running out of time. </p><p><strong>Instructions</strong></p><ul><li>Please answer the questions in English, or your test result will not count. </li><li>Do not click the Assessment Link until you are ready to take the assessment in full. Set aside 120 minutes in a quiet location where you will not be interrupted. </li><li>Ensure you have a reliable internet connection.</li><li>Please use one of the latest versions of Google Chrome, Firefox or Safari.</li><li>Languages available: Java, C, C++, C#, Python, Ruby, JavaScript, Golang, etc.</li></ul><p><strong>Assessment Link:</strong> <a href="{}">{}</a></p><p></br><br>Thank you and good luck!</p><p>Hao Zhang</p>'.format(name, interviewer_url, interviewer_url)
        mail.Send()
        print('Successfully sent email to {}'.format(name))