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

loc = ("SBM_Inverse.xls") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 

wr=Workbook()
sheet1=wr.add_sheet('Sheet 1')
sh2=wr.add_sheet('Sheet 2')
for i in range(n):
    sheet1.write(i+1,0,sheet.cell_value(i+1,0))
    
for i in range(n):
    sh2.write(i+1,0,sheet.cell_value(i+1,0))

sh2.write(0,1,'Beta')  

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
    E.append(prob.objective.value())

    for i in range(m):
        sheet1.write(k+1,i+2,value(s_negative[i]))
        
    for i in range(s):
        sheet1.write(k+1,i+m+2,value(s_positive[i]))


#Inverse SBM DEA    
sheet2= wb.sheet_by_index(1)
Z=[]

for i in range(m):
    D=[]
    for j in range(n):
        D.append(sheet2.cell_value(j+1,i+1))
    Z.append(D)
    


for k in range(n):
    o=k+1 #The dmu which you want to evaluate       
    s_negative_i=[0]*m
    lamda_i=[0]*n
    prob_inverse=LpProblem("Increase in Output",LpMaximize)
    for i in range(m):
        s_negative_i[i]=LpVariable("Slacknegative"+str(i),0)
    for i in range(n):
        lamda_i[i]=LpVariable("lamda"+str(i),0)
    
    beta=LpVariable("beta"+str(i),0)

    s_negative_i=np.array(s_negative_i)
    lamda_i=np.array(lamda_i)
    X=np.array(X)
    Y=np.array(Y)

    prob_inverse+= beta,"objective"
    
    dummy1=0
    for i in range(m):
        dummy1+=(1/m) - s_negative_i[i]*(1/((X[i][o-1]+Z[i][o-1])*m))  
    prob_inverse+= dummy1==E[o-1],"constraint"

    lamda_i.transpose()
    lamdaX_i=np.matmul(X,lamda_i)
    lamdaY_i=np.matmul(Y,lamda_i)

    for i in range(m):
        prob_inverse+= lamdaX_i[i]== (X[i][o-1]+Z[i][o-1]) - s_negative_i[i],str(i)+"Ayush"
            
        
    for i in range(s):
        prob_inverse+= lamdaY_i[i]== Y[i][o-1]+beta,str(i)+"Gupta"


    prob_inverse.solve()
    sh2.write(k+1,1,(value(beta)/Y[i][o-1])*100)

wr.save('Output_Increase(Percent).xls')