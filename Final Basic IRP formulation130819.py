# -*- coding: utf-8 -*-
"""
Created on Thu Aug  8 00:43:17 2019

@author: Karnika
"""

# -*- coding: utf-8 -*-
"""
Created on Mon Aug  5 23:38:00 2019

@author: Karnika
"""
from gurobipy import*
import os
import xlrd
from itertools import*
from scipy import spatial
from sklearn.metrics.pairwise import euclidean_distances
import math
from itertools import*

book = xlrd.open_workbook(os.path.join("DATA.xlsx"))

N=[]
cij={}
h={}
U={}
K=['V1']
T=[1,2,3]
r={} 
Iit_1={}
x={}
y={}
x1={}
x2={}
sh = book.sheet_by_name("abs1n10")


Q=sh.cell_value(1,8)            #Fleet Capacity

i = 1
while True:
    try:
        sp = sh.cell_value(i,0)
        N.append(sp)
        x[sp]=sh.cell_value(i,1)
        y[sp]=sh.cell_value(i,2)
        Iit_1[sp]=sh.cell_value(i,3)
        U[sp]=sh.cell_value(i,4)
        r[sp]=sh.cell_value(i,6)
        h[sp]=sh.cell_value(i,7)
        i = i + 1
    except IndexError:
        break
def calculate_dist(x1, x2):
    eudistance = spatial.distance.euclidean(x1, x2)    
    return(eudistance)    
for i in N:
    x1[i]=(x[i],y[i])
    for j in N:
        x2[j]=(x[j],y[j])
        cij[i,j]=int(round(calculate_dist(x1[i],x2[j])))
        
m=Model("Basic Model Of Inventory Routing Problem")

m.modelSense=GRB.MINIMIZE

yijkt=m.addVars(N,N,K,T,vtype=GRB.INTEGER,name='Y_ijkt')
Iit = m.addVars(N,T,vtype=GRB.CONTINUOUS,name='I_it')
qikt = m.addVars(N,K,T,vtype=GRB.CONTINUOUS,name='q_ikt')
Uikt=m.addVars(N,K,T,vtype=GRB.CONTINUOUS,name='U_ikt')
##Constraint 12:-
zikt= m.addVars(N,K,T,vtype=GRB.BINARY,name='Z_ikt')
##Constraint 1:-
m.setObjective(sum(h[i]*Iit_1[i] for i in N)+
               sum(h[i]*Iit[i,t] for i in N for t in T ) +
               sum(cij[i,j]*yijkt[i,j,k,t] for i in N for j in N for k in K for t in T ))
####for initialization 
####to provide initial inventory level of supplier and customer

######initialization complete
#####Constraint 2:-        

for t in T:
    if t == 1:
        m.addConstr((Iit_1[1] + r[1]- sum(qikt[i,k,t] for i in N if i != 1 for k in K)) == Iit[1,t] )
    else:
        m.addConstr((Iit[1,t-1] + r[1]- sum(qikt[i,k,t] for i in N if i != 1 for k in K)) == Iit[1,t] )

##Constraint 3:-
for i in N:
    if i!=1:
         for t in T:
            if t == 1:
                m.addConstr(Iit_1[i] + sum(qikt[i,k,t] for k in K) - r[i] == Iit[i,t] )
            else:
                m.addConstr(Iit[i,t-1]+sum(qikt[i,k,t] for k in K) - r[i] == Iit[i,t])
                
##Constraint 4:-
for i in N:
    for t in T:
        m.addConstr( Iit[i,t] >= 0)
                
for i in N:
    if i > 1:
        for t in T:
            m.addConstr(Iit[i,t]<=U[i])
        
##Constraint 5:-
for i in N:
    if i != 1:
        for k in K:
            for t in T:
                if t == 1:
                    m.addConstr(sum(qikt[i,k,t] for k in K) >= U[i]*sum(zikt[i,k,t] for k in K)-Iit_1[i]) 
                else:
                    m.addConstr(sum(qikt[i,k,t] for k in K)>= U[i]*sum(zikt[i,k,t] for k in K)-Iit[i,t-1])
####Constraint 6:-        
for i in N:
    if i!=1 :
        for t in T:
            if t == 1:
                m.addConstr(sum(qikt[i,k,t] for k in K)<= U[i] - Iit_1[i])
            else:
                m.addConstr(sum(qikt[i,k,t] for k in K)<= U[i] - Iit[i,t-1])
             
##Constraint 7:-
for i in N:
    if i != 1:
        for k in K:
            for t in T:
                m.addConstr(qikt[i,k,t]<=U[i]*zikt[i,k,t])

####Constraint 8:-                    
###constraints: These constraints guarantee that for each time t e ET,
###a feasible route is determined to visit all retailers served at time t.
###They can be for mulated as follows: 
# ##   (a) If at least one retailer s e N' is visited at time t,
###    then the route traveled at time t has to "visit" the supplier.
# ##   Let zot be a binary variable equal to one if the supplier is visited at time t 
###    and zero otherwise; then,
for k in K:
     for t in T:
         m.addConstr(sum(qikt[i,k,t] for i in N if i!=1 ) <= Q*zikt[1,k,t] )

##Constraint 9:-
for i in N:
    if i !=1:
        for t in T:
            m.addConstr(sum(zikt[i,k,t] for k in K)<=1)

#####Constraint 10:-                
for i in N:
    for k in K:
        for t in T:
            m.addConstr(sum(yijkt[i,j,k,t] for j in N  if j!=i)==zikt[i,k,t]) 
            m.addConstr(sum(yijkt[j,i,k,t] for j in N  if i!=j )==zikt[i,k,t])

##Constraint 13:-
for i in N:
    if i!=1:
        for k in K:
            for t in T:
                m.addConstr(qikt[i,k,t]>=0)                                   
####Constraint 14 & 15:-
for j in N:
    for k in K:
        for t in T:
            for i in N:
                if i == 1:
                    m.addConstr(yijkt[i,j,k,t] <= 2)
                else:
                    m.addConstr(yijkt[i,j,k,t] <= 1)
#####Subtour Elimination Constraint:-
for i in N:
    if i!=1:
        for t in T:
            for k in K:
                for j in N:
                    if i!=j:
                        m.addConstr((Uikt[i,k,t]-Uikt[j,k,t] + Q*yijkt[i,j,k,t])<= Q - qikt[i,k,t])
for i in N:
    if i!=1:
        for t in T:
            for k in K:
                m.addConstr(Uikt[i,k,t]<=Q) and m.addConstr(Uikt[i,k,t]>=qikt[i,k,t])
#N_dash=N[1:]
#SS=[]
#for i in N_dash:
#    i=int(i)
#    b=list(permutations(N_dash,i))
#    SS=SS+b                         
################################################################################                   
##Constraint 11:-
#for S in SS:
#    for mm in S:
#        for k in K:
#            for t in T:
#                m.addConstr(sum(yijkt[i,j,k,t] for i in S for j in S if i!=j)<=sum(zikt[i,k,t] for i in S)-zikt[mm,k,t], "sub-tour") 

m.write('MTZ.lp')
m.optimize()

for v in m.getVars():
    if v.x > 0.01:
        print(v.varName, v.x)
print('Objective:',m.objVal)
