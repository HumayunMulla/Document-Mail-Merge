#   Developer:  Humayun Mulla
#   Program:    A program to mail merge by reading sender's address from an excel file along with the contents of the mail body. 
#               It sends an email from outlook application. 
#   Date:       02/21/2020

import os
import datetime
import win32com.client as win32
import xlrd

print "This is working!"

# Win32 API for Outlook Application
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

receipients = []

# Sending email function
def send_email():  
    mail.To = "hmulla1@binghamton.edu"
    mail.Subject = "Testing Mail, please ignore"
    mail.Body = "TESTING EMAIL BODY"
    #mail.Send() # send email
    #print "Successfully email sent!"

# calling the send email function
# send_email()   

# Opening excel or *.xlsx file for reading recipient mail details
book = xlrd.open_workbook("recipient_details.xls")
sheet = book.sheet_by_index(0)

i = 1
while i <= 2: 
    for cell in sheet.row(i):
        
        receipients.append(cell.value)
        #send_email()
        #print cell.value
    # empty the list 
    
    i = i + 1

# Get System Time and Date
system_date = datetime.datetime.now().strftime('%d%m%Y'+"_"+'%H%M%S')
# print system_date 
# system_date variable is used later in the program for naming the file.

# Generate log for every email sent