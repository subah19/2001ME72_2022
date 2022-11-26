#Help
def scorecard():
	pass
from datetime import datetime
start_time = datetime.now()

###Code
####importing required libraries
import openpyxl
import pandas as pd
import os

#reading input files
india_inn = open("india_inns2.txt","r+") #india batting
pak_inn = open("pak_inns1.txt","r+") #pakistan batting
Playing_teams = open("teams.txt","r+")
input_team = Playing_teams.readlines()

pak_team = input_team[0]
pak_cricketers = pak_team[23:-1:].split(",")

ind_team = input_team[2]
ind_cricketers = ind_team[20:-1:].split(",")


lst_ind=india_inn.readlines() #124
for i in lst_ind:
    if i=='\n':
        lst_ind.remove(i)
      

lst_pak=pak_inn.readlines() #123
for i in lst_pak:
    if i=='\n':
        lst_pak.remove(i)

wb = openpyxl.Workbook()
sheet = wb.active

# batting [runs,ball,4s,6s,sr]
# bowling [over,medan,runs,Wickets, NB, WD, ECO]
#declaring required variables
Ind_out_count=0
FOW_pak=0
Pak_out_count={}
ind_bowlers={}
ind_bats={}

pak_batsman={}
pak_bowlers={}
pak_byes=0
Pak_bowlers_runs=0

########Pakistan Innings####################
for l in lst_pak:
    x=l.index(".")
    Pak_inn_overs=l[0:x+2]
    temp=l[x+2::].split(",")
    c_ball=temp[0].split("to") #0 2
    
    if f"{c_ball[0].strip()}" not in ind_bowlers.keys() :
        ind_bowlers[f"{c_ball[0].strip()}"]=[1,0,0,0,0,0,0]   #[over0,medan1,runs2,Wickets3, NB4, WD5, ECO6]
    elif "wide" in temp[1]:
        pass
    elif "byes" in temp[1]:                 #defining scores of byes
        if "FOUR" in temp[2]:
            pak_byes+=4
            ind_bowlers[f"{c_ball[0].strip()}"][0]+=1
        elif "1 run" in temp[2]:
            pak_byes+=1
            ind_bowlers[f"{c_ball[0].strip()}"][0]+=1
        elif "2 runs" in temp[2]:
            pak_byes+=2
            ind_bowlers[f"{c_ball[0].strip()}"][0]+=1
        elif "3 runs" in temp[2]:
            pak_byes+=3
            ind_bowlers[f"{c_ball[0].strip()}"][0]+=1
        elif "4 runs" in temp[2]:
            pak_byes+=4
            ind_bowlers[f"{c_ball[0].strip()}"][0]+=1
        elif "5 runs" in temp[2]:
            pak_byes+=5
            ind_bowlers[f"{c_ball[0].strip()}"][0]+=1

    else:
        ind_bowlers[f"{c_ball[0].strip()}"][0]+=1
    
    if f"{c_ball[1].strip()}" not in pak_batsman.keys() and temp[1]!="wide":
        pak_batsman[f"{c_ball[1].strip()}"]=[0,1,0,0,0] #[runs,ball,4s,6s,sr]
    elif "wide" in temp[1] :
        pass
    else:
        pak_batsman[f"{c_ball[1].strip()}"][1]+=1
    

    if "out" in temp[1]:                           #updating scoresheet when out
        ind_bowlers[f"{c_ball[0].strip()}"][3]+=1
        if "Bowled" in temp[1].split("!!")[0]:
            Pak_out_count[f"{c_ball[1].strip()}"]=("b" + c_ball[0])
        elif "Caught" in temp[1].split("!!")[0]:
            w=(temp[1].split("!!")[0]).split("by")
            Pak_out_count[f"{c_ball[1].strip()}"]=("c" + w[1] +" b " + c_ball[0])
        elif "Lbw" in temp[1].split("!!")[0]:
            Pak_out_count[f"{c_ball[1].strip()}"]=("lbw  b "+c_ball[0])

    
       #updating scoresheet when run made by bat
    if "no run" in temp[1] or "out" in temp[1] :
        ind_bowlers[f"{c_ball[0].strip()}"][2]+=0
        pak_batsman[f"{c_ball[1].strip()}"][0]+=0
    elif "1 run" in temp[1]:
        ind_bowlers[f"{c_ball[0].strip()}"][2]+=1
        pak_batsman[f"{c_ball[1].strip()}"][0]+=1
    elif "2 runs" in temp[1]:
        ind_bowlers[f"{c_ball[0].strip()}"][2]+=2
        pak_batsman[f"{c_ball[1].strip()}"][0]+=2
    elif "3 runs" in temp[1]:
        ind_bowlers[f"{c_ball[0].strip()}"][2]+=3
        pak_batsman[f"{c_ball[1].strip()}"][0]+=3
    elif "4 runs" in temp[1]:
        ind_bowlers[f"{c_ball[0].strip()}"][2]+=4
        pak_batsman[f"{c_ball[1].strip()}"][0]+=4
    elif "FOUR" in temp[1]:
        ind_bowlers[f"{c_ball[0].strip()}"][2]+=4
        pak_batsman[f"{c_ball[1].strip()}"][0]+=4
        pak_batsman[f"{c_ball[1].strip()}"][2]+=1
    elif "SIX" in temp[1]:
        ind_bowlers[f"{c_ball[0].strip()}"][2]+=6
        pak_batsman[f"{c_ball[1].strip()}"][0]+=6
        pak_batsman[f"{c_ball[1].strip()}"][3]+=1
    elif "wide" in temp[1]:                   #updating scoresheet when wide 
        if "wides" in temp[1]:
            # print(temp[1][1])
            ind_bowlers[f"{c_ball[0].strip()}"][2]+=int(temp[1][1])
            ind_bowlers[f"{c_ball[0].strip()}"][5]+=int(temp[1][1])
        else:
            ind_bowlers[f"{c_ball[0].strip()}"][2]+=1
            ind_bowlers[f"{c_ball[0].strip()}"][5]+=1

for val in pak_batsman.values():
    val[-1]=round((val[0]/val[1])*100 , 2)
