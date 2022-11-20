from datetime import datetime
from unicodedata import name
start_time = datetime.now()

def attendance_report():
###Code
#importing required libraries
 import csv 
 import os  
 import numpy as np 
 os.system('cls')
 RN=[] #list of students' roll number
 Name=[] #list of students' name

 with open('input_registered_students.csv', 'r') as file: 
  reader = csv.reader(file)
  n=0
  for row in reader:
   if n!=0:
    RN.append(row[0]) 
    Name.append(row[1])   
   n=n+1 
  n=n-1
 dat =["28-07","01-08","04-08","08-08","11-08","15-08","18-08","22-08","25-08","29-08","01-09","05-09","08-09","12-09","15-09","26-09","29-09"]
 # date on which class was conducted
 s=len(dat) # total dates


 with open('input_attendance.csv', 'r') as file: #opening attendance file
  reader = csv.reader(file) 
  
  TD=[] 
  SD=[] 
  st=[] 
  for x in RN:
   SD=[] # Initializing Stud data
   for j in dat:
    st=[] 
    RA=0 
    DA=0  
    FA=0  
    with open('input_attendance.csv', 'r') as file: #opening attendance file again
     reader = csv.reader(file)
     for row in reader:
      if j==row[0][0:5] and row[1][0:8]==x and (row[0][11:13]=="14" or row[0][11:16]=="15:00" ) and RA==0 : 
       RA=RA+1 # Updating real attendance of student
      elif j==row[0][0:5] and row[1][0:8]==x and (row[0][11:13]=="14" or row[0][11:16]=="15:00") and RA>0 : 
       DA=DA+1 # Updating Duplicate attendance of student 
      elif j==row[0][0:5] and row[1][0:8]==x :
       FA=FA+1 
     st.append(RA+DA+FA)   
     st.append(RA)  
     st.append(DA)   
     st.append(FA)  
     if RA==0:
      st.append(1)
     else:
      st.append(0) 
    SD.append(st)
   TD.append(SD)
 
 if os.path.exists("output"):
  for f in os.listdir("output"):
    os.remove(os.path.join("output",f)) # removing all the prebuild files in output folder

 os.chdir("output")
 from openpyxl import Workbook
 for i in range(0,n): 
  book=Workbook()
  sheet= book.active    
  rows=[] # Making of list of rows of a particular student of all dates
  rows.append(["Date","Roll No.","Name","Total attendance count","Real","Duplicate","Invalid","absent"])
  rows.append(["",RN[i],Name[i],"","","","",""])
  for q in range(0,s):
   rows.append([dat[q],"","",TD[i][q][0],TD[i][q][1],TD[i][q][2],TD[i][q][3],TD[i][q][4]]) # Appending all types of Attendance in row
  for w in rows:
   sheet.append(w)
  book.save( RN[i] + ".xlsx") 
   
  dic={0:"A",1:"P"} # Dictionary for present and absent
  book=Workbook()
  sheet= book.active    
  rows=[]
  z=["Roll No.","Name"] # Initializing 1st row
  for i in dat: # Initializing 1st row
   z.append(i)
  z.append("Total Lecture taken") 
  z.append("Total Real")
  z.append("% Attendance")
  rows.append(z) 

  z=["(Sorted by roll no.)","","Atleast one real P"] # Initializing 2nd row
  for i in range(0,s-1): # using this loop we are making 2nd row
   z.append("")
  z.append("(=Total Mon+Thur dynamic count")
  z.append("")
  z.append("Real/Actual Lecture taken")
  rows.append(z) 

  for i in range(0,n):  # using this loop i am making full data of all the students
   z=[RN[i],Name[i]]
   SD=0
   for q in range(0,s):
    z.append(dic[TD[i][q][1]]) #total lecture taken
    if dic[TD[i][q][1]]=="P":
     SD=SD+1
   z.append(s)
   z.append(SD) 
   l=(SD/s)*100
   z.append("{:.2f}".format(l))
   rows.append(z)

  for w in rows:
   sheet.append(w)

   book.save( "Attendance_report_consolidated" + ".xlsx")  # Saving full attendance report
  



attendance_report()
# importing required libraries for sending mail

    


            
def send_mail(): #defining function to send mail
    import smtplib
    import email.encoders
    import base64

    from email.encoders import encode_base64
    from getpass import getpass
    from smtplib import SMTP
    
    from email import encoders
    from email.mime.base import MIMEBase
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from mimetypes import guess_type
    fromaddr = input("Enter Mail Id: ") 
    toaddr = "cs3842022@gmail.com"  
    Password_ = input("Enter Password: ")  

   
    msg = MIMEMultipart()    
    msg['From'] = fromaddr    
    msg['To'] = toaddr    
    msg['Subject'] = " Attendance report"   
    body = "Dear Sir,\n\nPlease find attachment.\n\nThanks and Regards\nShubham\n2001ME72" 
    msg.attach(MIMEText(body, 'plain')) 

    filename = 'attendance_report_consolidated.xlsx'  
    attachment = open("attendance_report_consolidated.xlsx", "rb") 

    p = MIMEBase('application', 'octet-stream')  

    p.set_payload((attachment).read())    

    encoders.encode_base64(p)            

    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(p)    

    s = smtplib.SMTP('smtp.gmail.com', 587)     

    s.starttls()     

    s.login(fromaddr, Password_)  

    text = msg.as_string()      

    s.sendmail(fromaddr, toaddr, text)  
    s.quit()



send_mail()   




#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))