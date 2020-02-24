#   Developer:  Humayun Mulla
#   Program:    A program to mail merge by reading sender's address from an excel file along with the contents of the mail body. 
#               It sends an email from outlook application. 
#   Date:       02/21/2020
import os
import datetime
import win32com.client as win32
import xlrd

# Certain details are abstracted and defined in config.ini file
# Config Parser
try:
    from configparser import ConfigParser
except ImportError:
    from ConfigParser import ConfigParser  # ver. < 3.0
# instantiate
config = ConfigParser()
# parse existing file
config.read('config.ini')
mail_subject = config.get('system', 'Subject')
BodyText1 = config.get('system', 'BodyText1')
BodyText2 = config.get('system', 'BodyText2')
BodyText3 = config.get('system', 'BodyText3')
BodyText4 = config.get('system', 'BodyText4')
Signature = config.get('system', 'Signature')
Designation = config.get('system', 'Designation')

# print "This is working!"



# Sending email function
def send_email(receipient_name, to_address, body_content1, body_content2):  
    # Win32 API for Outlook Application
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    body_detailed = str(body_content2).split(', ')
    # print body_detailed
    i = 0
    mail_sublist = ""
    sizeofList = len(body_detailed)
    while i < sizeofList:
        mail_sublist += str(body_content1) + "\t" + body_detailed[i] + "\n"
        i += 1
    # print mail_sublist
    # print receipient_name
    mail.To = to_address
    mail.Subject = mail_subject
    # print mail_subject
    mail_body = "Hello " + receipient_name + ",\n"
    mail_body = mail_body + "\n" + BodyText1 + "\n\n" + BodyText2 + "\n\n" + mail_sublist + "\n" + BodyText3 + "\n\n"BodyText4"\n\n" + Signature +"\n" + Designation
    
    mail.Body = mail_body
    # print mail.Body
    mail.Send() # send email
    print "Successfully email sent!"
    
    

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
receipient_body_part1 = []
receipient_body_part2 = []

for row_cursor in range(1,total_rows):
    # excel_data = worksheet.cell(row_cursor,1).value
    # Receipient Name
    receipient_name.append(worksheet.cell(row_cursor,2).value)
    # Receipient Email Address
    receipient_email.append(worksheet.cell(row_cursor,3).value)
    # Receipient Body Part-1
    # receipient_body_part1 = str(worksheet.cell(row_cursor,4).value)
    receipient_body_part1.append(str(worksheet.cell(row_cursor,4).value))
    
    # Receipient Body Part-2
    col_index = 5 # customized details start from this location
    while col_index < 15:
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
    send_email(receipient_name[i], receipient_email[i], receipient_body_part1[i], receipient_body_part2[i])

    i += 1

# print receipient_body_part1
#print receipient_body_part2[0]


# Get System Time and Date
system_date = datetime.datetime.now().strftime('%d%m%Y'+"_"+'%H%M%S')
# print system_date 
# system_date variable is used later in the program for naming the file.

# Generate log for every email sent

