# -*- coding: utf-8 -*-
"""
Created on Wed Dec  7 10:35:56 2022

@author: E1448105
"""
from Surveyemail import findtopn
from openpyxl import load_workbook
import win32com.client as win32

book = load_workbook("./Student Awards.xlsx")
s = book["Sheet1"]
body = s["A1"].value

def add_postfix(number):

    try :
        tens = number[-2]
    except IndexError:
        tens = ""
    ones = number[-1]

    if len(number) >1 and tens =="1":
        number += "th"
    elif ones == "1":
        number += "st"
    elif ones == "2":
        number += "nd"
    elif ones == "3":
        number+= "rd"
    else:
        number += "th"
    return number

time30,visit30 = findtopn(25)
merch_rank = -1
for i,student in enumerate(time30):
    rank = str(i+1)
    if i%10 == 0:
        merch_rank +=1

    rank = add_postfix(rank)


    merch_list = ["MCC T-shirt","MCC tumbler","MCC lanyard"]
    merch = merch_list[merch_rank]
    student.subject = "Congrats from the Student Success Center!"
    student.body = body.replace("[NAME]",student.FN).replace("[RANK]",rank).replace("[VISITS]",str(student.visits)).replace("[HOURSMINUTES]", student.hoursminutestext()).replace("[MERCH]",merch)

def send_email(students):
    for student in students:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)  # creates an email object
        mail.SentOnBehalfOfName = "lvlr.studentsuccess@mcckc.edu"
        mail.To = student.email
        mail.Subject = student.subject
        mail.Body = student.body
        mail.Send()
        print(f"Email sent to {student.FN}")