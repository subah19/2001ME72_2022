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

