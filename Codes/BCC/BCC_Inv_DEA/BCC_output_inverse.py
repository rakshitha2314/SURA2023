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
E=[]           #This is the matrix containing efficiency
loc = ("BCC_inverse.xls") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
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
        
for i in range(s):
    C=[]
    for j in range(n):
        C.append(sheet.cell_value(j+1,i+m+1))
    Y.append(C)

X=np.array(X)
Y=np.array(Y)


#DEA code

for k in range(n):
    o=k+1 #The dmu which you want to evaluate       

    lamda=[0]*n
    prob=LpProblem("Efficiency Ranking",LpMaximize)

    for i in range(n):
        lamda[i]=LpVariable("lamda"+str(i),0)
    t=LpVariable("Additive",0)

    lamda=np.array(lamda)

    prob+= t ,"objective"

    lamda.transpose()
    lamdaX=np.matmul(X,lamda)
    lamdaY=np.matmul(Y,lamda)

    for i in range(m):
        prob+= lamdaX[i]<= X[i][o-1],str(i)+"Ayush"
        
    for i in range(s):
        prob+= lamdaY[i]>= t*Y[i][o-1],str(i)+"Gupta"
    
    dummy1=0
    for i in range(n):
        dummy1+=lamda[i]
    
    prob+= dummy1 == 1 , "constraint"


    prob.solve()
    sheet1.write(k+1,1,prob.objective.value())
    E.append(prob.objective.value())
    

#Inverse DEA code

sheet2= wb.sheet_by_index(1)
Z=[]

for i in range(m):
    D=[]
    for j in range(n):
        D.append(sheet2.cell_value(j+1,i+1))
    Z.append(D)


for i in range(s):
    sheet1.write(0,i+2,"beta"+str(i))

for k in range(n):
    o=k+1
    weights=[]
    for i in range(s):
        weights.append(1/s)
        
    lamda=[0]*n
    beta=[0]*s #Here beta stands for increment in output for a particular increment in input
    prob_inverse=LpProblem("Max increase in Outputs",LpMaximize)
    
    for i in range(n):
        lamda[i]=LpVariable("lamda"+str(i),0) #All lamda are greater than 0
    
    for i in range(s):
        beta[i]=LpVariable("beta"+str(i),0) #All beta are greater than 0
    
    dummy1=0
    for i in range(n):
        dummy1+=lamda[i]
    
    prob_inverse+= dummy1 == 1 , "constraint" #Summation of lamda is equal to 1
    
    dummy2=0
    for i in range(s):
        dummy2+=beta[i]*weights[i]
    
    prob_inverse+= dummy2 ,"objective"
    
    
    lamda=np.array(lamda)
    
    lamda.transpose()
    lamdaX=np.matmul(X,lamda)
    lamdaY=np.matmul(Y,lamda)
    
    for i in range(m):
        prob_inverse+= lamdaX[i]<= X[i][o-1]+Z[i][o-1],str(i)+"Ayush"
        
    for i in range(s):
        prob_inverse+= lamdaY[i]>= (Y[i][o-1]+beta[i])*E[o-1],str(i)+"Gupta"
        
    
    prob_inverse.solve()
    
    for i in range(s):
        sheet1.write(k+1,i+2,(value(beta[i])/Y[i][o-1])*100)
        
wr.save('Output_increase_BCC.xls')
        
        
    
    
        
    
    
    



        





