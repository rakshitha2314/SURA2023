from pulp import *
import xlrd
import xlwt
from xlwt import Workbook
import numpy as np
m=int(input()) #This is the number of inputs we want
s=int(input()) #This is the number of outputs we want
n=int(input()) #This is the total number of DMUs
X=[]           #This is the matrix containing the inputs
Y=[]           #This is the matrix containing the outputs 

loc = ("SURA_Final_Submit.xls") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 

wr=Workbook()
sheet1=wr.add_sheet('Sheet 1')
for i in range(n):
    sheet1.write(i+1,0,sheet.cell_value(i+1,0))
sheet1.write(0,1,'Efficiency')
for i in range(m):
    sheet1.write(0,i+2,"s_negative"+str(i))
for i in range(s):
    sheet1.write(0,i+m+2,"s_positive"+str(i))

for i in range(m):
    D=[]
    for j in range(n):
        D.append(sheet.cell_value(j+1,i+1))
    X.append(D)
        
for i in range(s):
    C=[]
    for j in range(n):
        C.append(sheet.cell_value(j+1,i+m+1))
    Y.append(C)


for k in range(n):
    o=k+1 #The dmu which you want to evaluate       
    s_positive=[0]*s
    s_negative=[0]*m
    lamda=[0]*n
    prob=LpProblem("Efficiency Ranking",LpMinimize)
    for i in range(s):
        s_positive[i]=LpVariable("Slackpositive"+str(i),0)
    for i in range(m):
        s_negative[i]=LpVariable("Slacknegative"+str(i),0)
    for i in range(n):
        lamda[i]=LpVariable("lamda"+str(i),0)
    t=LpVariable("Additive",0)

    s_positive=np.array(s_positive)
    s_negative=np.array(s_negative)
    lamda=np.array(lamda)
    X=np.array(X)
    Y=np.array(Y)

    dummy1=0
    for i in range(m):
        dummy1+=t*(1/m) - s_negative[i]*(1/(X[i][o-1]*m))  
    prob+= dummy1 ,"objective"

    dummy2=0
    for i in range(s):
        dummy2+=t*(1/s) + s_positive*(1/(Y[i][o-1]*s))

    prob+= dummy2 == 1 , "constraint"

    lamda.transpose()
    lamdaX=np.matmul(X,lamda)
    lamdaY=np.matmul(Y,lamda)

    for i in range(m):
        prob+= lamdaX[i]== t*X[i][o-1] - s_negative[i],str(i)+"Ayush"
        
    for i in range(s):
        prob+= lamdaY[i]== t*Y[i][o-1] + s_positive[i],str(i)+"Gupta"


    prob.solve()
    sheet1.write(k+1,1,prob.objective.value())

    for i in range(m):
        sheet1.write(k+1,i+2,value(s_negative[i]))
        
    for i in range(s):
        sheet1.write(k+1,i+m+2,value(s_positive[i]))
    
wr.save('SBM_Final_Output.xls')


        





