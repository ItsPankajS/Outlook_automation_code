# -*- coding: utf-8 -*-


import win32com.client as client


outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To ='test@test.com' #Test Code Line 1
message.CC = 'test@test.com'
message.Subject ='Test Subject'
message.body = ''
message.body = 'Hi Ram,\n\nGood Afternoon,\n\
This is a test email.'

message.Send()
print('Email Sent to Business')




