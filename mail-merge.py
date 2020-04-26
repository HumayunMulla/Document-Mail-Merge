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
BodyText3A = config.get('system', 'BodyText3A')
BodyText3B = config.get('system', 'BodyText3B')
hyperlink = config.get('system', 'hyperlink')
hyperlinkName = config.get('system', 'hyperlinkName')
BodyText4 = config.get('system', 'BodyText4')
Signature = config.get('system', 'Signature')
Certification = config.get('system', 'Certification')

# print "This is working!"

# Get System Time and Date
system_date = datetime.datetime.now().strftime('%d%m%Y'+"_"+'%H%M%S')
# print system_date 
# system_date variable is used later in the program for naming the file.
fileOpen = open(system_date+".txt","w+")

# Generate log for every email sent
def generate_log(receipient_name, to_address):
    timestamp = datetime.datetime.now().strftime('%d/%m/%Y'+" "+'%H:%M:%S')
    fileOpen.write("Successfully email sent to " + receipient_name + " [" + to_address + "]" + " at " + timestamp + "\n")

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
        # mail_sublist += str(body_content1) + "\t" + body_detailed[i] + "\n"
        mail_sublist += str(body_content1) + ": \t\t" + body_detailed[i] + "<br>"
        i += 1
    # print mail_sublist
    # print receipient_name
    mail.To = to_address
    mail.Subject = mail_subject
    # print mail_subject
    # mail_body = "Hello " + receipient_name + ",\n"
    # mail_body = mail_body + "\n" + BodyText1 + "\n\n" + BodyText2 + "\n\n" + mail_sublist + "\n" + BodyText3 + "\n\n" + BodyText4 +"\n\n" + Signature +"\n" + Designation
    attachment = mail.Attachments.Add("C:\source\Document-Mail-Merge\signature.png", 0x5, 0, "photo")
    imageCid = "signature.png@123"
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imageCid)
    # print mail_sublist
    mail_body = '<html><body><p style="font: 14px arial, sans-serif;"> Hello&nbsp;'+ receipient_name + ',<br><br>' + BodyText1 + '<br><br>' + BodyText2 + '<br><br>'+ mail_sublist + '<br><br>' + BodyText3A + '&nbsp;<a href="'+hyperlink+'">'+hyperlinkName +'</a>&nbsp;' + BodyText3B + '<br><br>' + BodyText4 + '</p><p style="font: bold 14px calibri, sans-serif;">'+Signature+'&nbsp;<span style="font: normal 10px calibri, sans-serif;">'+ Certification +'</span><br><img src=\"cid:{0}\" height=25 width=200></p></body></html>'.format(imageCid)
    # mail.HTMLBody = '<html><body><p style="font: 14px arial, sans-serif;"> Hello '+ receipient_name + ',<br><br>' + BodyText1 + '<br><br>' + BodyText2 + '<br><br></p><textarea>' + mail_sublist + '</textarea></body></html>'
    mail.HTMLBody = mail_body

    # print mail.Body
    mail.Send() # send email
    print "Successfully email sent!"
    generate_log(receipient_name, to_address)
    
    

# calling the send email function
# send_email()   

# Opening excel or *.xlsx file for reading recipient mail details
workbook = xlrd.open_workbook('recipient_details.xls')
worksheet = workbook.sheet_by_name('Sheet1')
# find the total number of rows in the sheet
total_rows = worksheet.nrows
# print total_rows

from collections import defaultdict
contact_dict = defaultdict(list)

# excel upload
for row_cursor in range(1,total_rows):
    # Receipient Name is the Key 
    # Receipient Email Address
    contact_dict[worksheet.cell(row_cursor,2).value].append(worksheet.cell(row_cursor,3).value)
    # Receipient Body Part-1
    contact_dict[worksheet.cell(row_cursor,2).value].append(worksheet.cell(row_cursor,4).value)
    
    # Receipient Body Part-2
    col_index = 5 # customized details start from this location
    while col_index < 15:
        if col_index==5:
            excel_data = worksheet.cell(row_cursor,5).value
            if isinstance(excel_data, float):
                #print 'excel_data is a float!'
                excel_data = int(excel_data)

            if excel_data !="":
                #print excel_data
                # receipient_body_part2.append(excel_data)
                contact_dict[worksheet.cell(row_cursor,2).value].append(excel_data)
        else:                
            excel_data = worksheet.cell(row_cursor,col_index).value
            if isinstance(excel_data, float):
                #print 'excel_data is a float!'
                excel_data = int(excel_data)

            if excel_data!="":                
                #print excel_data
                # receipient_body_part2[row_cursor-1] = str(receipient_body_part2[row_cursor-1]) + ", " + str(excel_data) # conversion into string required
                contact_dict[worksheet.cell(row_cursor,2).value].append(excel_data)
        
        col_index += 1
    
# print contact_dict

for key, value in contact_dict.iteritems() :
    send_email( key, value[0], value[1], value[2])



