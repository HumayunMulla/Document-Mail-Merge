#   Developer:  Humayun Mulla
#   Program:    A program to mail merge by reading sender's address from an excel file along with the contents of the mail body. 
#               It sends an email from outlook application. 
#   Date:       02/21/2020
import os
import datetime
import win32com.client as win32
import xlrd

# print "This is working!"

# Win32 API for Outlook Application
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)


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
workbook = xlrd.open_workbook('recipient_details.xls')
worksheet = workbook.sheet_by_name('Sheet1')
# find the total number of rows in the sheet
total_rows = worksheet.nrows
# print total_rows

# variables used to send emails [receipient addesss, name, subject contents]
receipient_email = []
receipient_name = []
receipient_body_part1 = ""
receipient_body_part2 = []

for row_cursor in range(1,total_rows):
    # excel_data = worksheet.cell(row_cursor,1).value
    # Receipient Name
    receipient_name.append(worksheet.cell(row_cursor,2).value)
    # Receipient Email Address
    receipient_email.append(worksheet.cell(row_cursor,3).value)
    # Receipient Body Part-1
    receipient_body_part1 = worksheet.cell(row_cursor,4).value
    # Receipient Body Part-2
    col_index = 5 # customized details start from this location
    while col_index < 16:
        if col_index==5:
            excel_data = worksheet.cell(row_cursor,5).value
            if excel_data !="":
                receipient_body_part2.append(excel_data)
        else:                
            excel_data = worksheet.cell(row_cursor,col_index).value
            if excel_data!="":
                receipient_body_part2[row_cursor-1] = str(receipient_body_part2[row_cursor-1]) + ", " + str(excel_data) # conversion into string required
        col_index += 1
    #print receipient_body_part2[row_cursor-1]
    
    

# printing just to check if correct values are populating 
# print receipient_name[0]
# print receipient_email[0]
# iterate using while loop and print the details
i = 0
sizeofList = len(receipient_email)
while i < sizeofList:
    print receipient_name[i] + " " + receipient_email[i] + " " + receipient_body_part1 + " " + receipient_body_part2[i]
    i += 1

# print receipient_body_part1
#print receipient_body_part2[0]


# Get System Time and Date
system_date = datetime.datetime.now().strftime('%d%m%Y'+"_"+'%H%M%S')
# print system_date 
# system_date variable is used later in the program for naming the file.

# Generate log for every email sent