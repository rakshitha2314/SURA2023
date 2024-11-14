from pulp import *
import xlrd
import xlwt
from xlwt import Workbook
import numpy as np
m=int(input())         #This is the number of inputs we want
s=int(input())         #This is the number of outputs we want
n=int(input())         #This is the total number of DMUs
phi=float(input())     #This represents the confidence level
sigma=float(input())   #This represents the variance
X=[]                   #This is the matrix containing the mean of inputs
Y=[]                   #This is the matrix containing the mean of outputs
A=[]                   #This is the matrix containing the variance factor of inputs 
B=[]                   #This is the matrix containing the variance factor of inputs 
loc = ("Stochastic_DEA.xls") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet_var = wb.sheet_by_index(1) 
wr=Workbook()
sheet1=wr.add_sheet('Sheet 1')
for i in range(n):
    sheet1.write(i+1,0,sheet.cell_value(i+1,0))
sheet1.write(0,1,'Efficiency')
for i in range(m):
    D=[]
    for j in range(n):
        D.append(sheet.cell_value(j+1,i+1))
    X.append(D)
for i in range(m):
    D=[]
    for j in range(n):
        D.append(sheet_var.cell_value(j+1,i+1))
    A.append(D)     
for i in range(s):
    C=[]
    for j in range(n):
        C.append(sheet.cell_value(j+1,i+m+1))
    Y.append(C)  
for i in range(s):
    C=[]
    for j in range(n):
        C.append(sheet_var.cell_value(j+1,i+m+1))
    B.append(C)
for k in range(n):
    o=k+1 #The dmu which you want to evaluate       
    p_positive=[0]*m
    p_negative=[0]*m
    q_positive=[0]*m
    q_negative=[0]*m
    lamda=[0]*n
    prob=LpProblem("Efficiency Ranking",LpMinimize)
    for i in range(s):
        q_positive[i]=LpVariable("q_positive"+str(i),0)
    for i in range(s):
        q_negative[i]=LpVariable("q_negative"+str(i),0)
    for i in range(m):
        p_postive[i]=LpVariable("p_positive"+str(i),0)
    for i in range(m):
        p_negative[i]=LpVariable("p_negative"+str(i),0)
    for i in range(n):
        lamda[i]=LpVariable("lamda"+str(i),0)
    p_positive=np.array(p_positive)
    p_negative=np.array(p_negative)
    q_positive=np.array(q_positive)
    q_negative=np.array(q_negative)
    lamda=np.array(lamda)
    X=np.array(X)
    Y=np.array(Y)
    A=np.array(A)
    B=np.array(B)
    theta=LpVariable("Stochastic",0)
    prob+= theta ,"objective"
    lamda.transpose()
    lamdaX=np.matmul(X,lamda)
    lamdaY=np.matmul(Y,lamda)
    lamdaA=np.matmul(A,lamda)
    lamdaB=np.matmul(B,lamda)
    for i in range(m):
        prob+= lamdaX[i] - phi*sigma*(p_positive[i]+p_negative[i])<= theta*X[i][o-1],str(i)+"Mean_x"
    for i in range(m):
        prob+= lamdaA[i] - theta*A[i][o-1]== (p_positive[i]-p_negative[i]),str(i)+"Var_x"
    for i in range(s):
        prob+= lamdaY[i] + phi*sigma*(q_positive[i]+q_negative[i])>= Y[i][o-1],str(i)+"Mean_y"
    for i in range(s):
        prob+= lamdaB[i] - B[i][o-1]== (q_positive[i]-q_negative[i]),str(i)+"Var_y"
    prob.solve()
    sheet1.write(k+1,1,prob.objective.value())
wr.save('Stochastic_Efficiency.xls')


        





