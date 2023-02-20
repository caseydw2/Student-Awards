# -*- coding: utf-8 -*-
"""
Created on Wed Nov  9 19:42:52 2022

Make list of students to send survey to.

@author: Casey
"""

import pandas as pd
import matplotlib.pyplot as plt


class Student:
    def __init__(self,ID, FN, LN, period , visits):
        self.FN = FN
        self.LN = LN
        self.id = ID
        self.email = str(ID) + "@student.mcckc.edu"
        self.hr = int(period.split(":")[0])
        self.minu = int(period.split(":")[1])
        self.visits = visits


    def __str__(self):
        return (str(self.FN) + " " + str(self.LN) + " (ID#: S" + str(self.id) + ")")

    def addPeriod(self,period):
        hrs = int(period.split(":")[0])
        mins = int(period.split(":")[1])
        self.hr += hrs
        self.minu += mins

    def mintohr(self):
        hrs,mins = divmod(self.minu,60)
        self.hr += hrs
        self.minu = mins

    def hoursminutestext(self):

        text = str(self.hr) + " hour"
        if self.hr > 1:
            text += "s"
        if self.minu >= 15:
            text += " and " + str(self.minu) +" minutes"
        return text

def sl():

    Snum, Fname, Lname, Period, Visits = 0,1,3,7,8

    path = "./FallAccudemiaLogs.xls"
    sheet = "Body"
    bodysheet = pd.read_excel(path,sheet)
    bodysheet = bodysheet.to_numpy()

    students = pd.read_excel(path)
    students = students.to_numpy()

    s0 = students[0]
    stud = [Student(s0[Snum],s0[Fname],s0[Lname],s0[Period],s0[Visits])]
    most_visited = Student(s0[Snum],s0[Fname],s0[Lname],s0[Period],s0[Visits])
    longest_visited = Student(s0[Snum],s0[Fname],s0[Lname],s0[Period],s0[Visits])
    for i, student in enumerate(students):
        if i == 0:
            continue
        elif student[Snum] != students[i-1][Snum]:
            studi = Student(student[Snum],student[Fname],student[Lname],student[Period],student[Visits])
            stud.append(studi)
        else:
            stud[-1].addPeriod(student[Period])
            stud[-1].visits += student[Visits]
    for student in stud:
        student.mintohr()

        if student.hr > longest_visited.hr:
            longest_visited = student
        elif student.hr == longest_visited.hr:
            if student.minu > longest_visited.minu:
                longest_visited = student

        if student.visits > most_visited.visits:
            most_visited = student

        body = str(bodysheet[0][0]).replace("[NAME]",student.FN).replace("[VISITS]", str(student.visits)).replace("[HOURSMINUTES]", str(student.hoursminutestext()))
        student.body = body
        student.subject = "End of Semester Questionnaire"
    print(longest_visited, "visited:", longest_visited.hr ," : ", longest_visited.minu, " ",most_visited, "  visited ", most_visited.visits )
    return stud

def stuemail():
    stumail = []
    students = sl()
    for student in students:
        if student.visits >= 3:
            stumail.append(student)
    return stumail

def findtopn(n=30,total = False):
    Snum, Fname, Lname, Period, Visits = 0,1,3,7,8

    path = "./FallAccudemiaLogs.xls"

    students = pd.read_excel(path,"SummaryByCourse")
    students = students.to_numpy()

    s0 = students[0]
    stud = [Student(s0[Snum],s0[Fname],s0[Lname],s0[Period],s0[Visits])]
    for i, student in enumerate(students):
        if i == 0:
            continue
        elif student[Snum] != students[i-1][Snum]:
            studi = Student(student[Snum],student[Fname],student[Lname],student[Period],student[Visits])
            stud.append(studi)
        else:
            stud[-1].addPeriod(student[Period])
            stud[-1].visits += student[Visits]
    for student in stud:
        student.mintohr()
        student.totmin = student.hr * 60 + student.minu
    if total:
        n = len(stud)
    topntime = sorted(stud, key=lambda x: x.totmin, reverse = True)
    topnvisit = sorted(stud,key=lambda x: x.visits, reverse = True)
    return topntime[:n],topnvisit[:n]

def toppercent(x):
    tot = sum(x)
    toptenpercent = int(len(x)*0.1)
    toptot = sum(x[:toptenpercent])
    return toptot/tot

def printtopntime(n=30,total= False):
    x = findtopn(n,total)[0]
    for i,student in enumerate(x):
        print( "Number ", i+1 ,": ", student.FN, " ", student.LN, " with ", student.hr, " hours and ", student.minu, " minutes!")

def printtopnvisit(n=30,total= False):
    x = findtopn(n,total)[1]
    for i,student in enumerate(x):
        print( "Number ", i+1 ,": ", student.FN, " ", student.LN, " with ", student.visits, " visits!")

def graphtopntime(n=30,total= False):
    y = findtopn(n,total)[0]
    x = [student.totmin/60 for student in y]
    allx = [student.totmin for student in y]
    plt.figure()
    plt.plot(x)
    plt.show()
    print(toppercent(allx))

def graphtopnvisit(n=30,total= False):
    y = findtopn(n,total)[1]
    x = [student.visits for student in y]
    allx = [student.visits for student in y]
    plt.figure()
    plt.plot(x)
    plt.show()
    print(toppercent(allx))





