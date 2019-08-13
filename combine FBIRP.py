# -*- coding: utf-8 -*-
"""
Created on Mon Aug 12 21:17:06 2019

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
import xlsxwriter

def IRP(DATA):

    book = xlrd.open_workbook(os.path.join("DATA.xlsx"))
    
    N=[]            #its a set of supplier and Customer,supplier at 1 and remaining are the Customers
    cij={}          #it represents cost of transportation to travel between locations i and j at non-negative cost cij 
    h={}            # per period unit inventory holding cost hi
    U={}            #A maximum inventory level Ui is also associated with each customer i belongs to N_Dash
    K=['V1']        #No of Fleets now only one vehicle is used 'V1'
    T=[1,2,3]       #Let T denote the set of time periods
    r={}            #rit units are consumed at customer i belongs to N_dash 
    Iit_1={}        #an initial inventory level Ii0 are associated with each node
    x={}            #x is set of  x Co-ordinate of Nodes belongs to N
    y={}            #y is set of  y Co-ordinate of Nodes belongs to N
    x1={}
    x2={}
    
    sh = book.sheet_by_name(DATA)
    
    
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
    #####  N_Dash:-
    N_dash=N[1:]
    ####Calculation of cij[i,j] matrix:
    
    def calculate_dist(x1, x2):
        eudistance = spatial.distance.euclidean(x1, x2)    
        return(eudistance)    
    for i in N:
        x1[i]=(x[i],y[i])
        for j in N:
            x2[j]=(x[j],y[j])
            cij[i,j]=int(round(calculate_dist(x1[i],x2[j])))
            
                                ##################################        
                        ########### " Model Formulation " #############
                                ##################################
      
    m=Model("Basic Model Of Inventory Routing Problem")
    
    m.modelSense=GRB.MINIMIZE
    
    yijkt=m.addVars(N,N,K,T,vtype=GRB.INTEGER,name='Y_ijkt')    #the variable yijkt represents the number of time
                                                                #the edge (i,j) belongs to E is traversed 
                                                                #by vehicle k in the time period t
    Iit = m.addVars(N,T,vtype=GRB.CONTINUOUS,name='I_it')       #used to indicate the inventory level at Node i belongs to N
                                                                #at the end of time period t
    qikt = m.addVars(N,K,T,vtype=GRB.CONTINUOUS,name='q_ikt')   #it represents the quantity delivered to customer i
                                                                #belongs to N_Dash
    Uikt=m.addVars(N,K,T,vtype=GRB.CONTINUOUS,name='U_ikt')
    ##Constraint 12:-
    zikt= m.addVars(N,K,T,vtype=GRB.BINARY,name='Z_ikt')        #its a binary variable 1 if node i belongs to N is visited
                                                                #at time period t belongs T by vehicle v
    ##Constraint 1:-
    
    #####The objective function (1) calls for the minimization of the total operational cost, that is the sum
    #####of the inventory costs at the depot, inventory costs at the customers, and costs of the routes over
    #####the time horizon.
    m.setObjective(sum(h[i]*Iit_1[i] for i in N)+
                   sum(h[i]*Iit[i,t] for i in N for t in T ) +
                   sum(cij[i,j]*yijkt[i,j,k,t] for i in N for j in N for k in K for t in T ))
    
    #####Constraint 2:-   
    #####Constraints (2)–(4) determine the evolution of the inventory level over time and
    #####force, the absence of stockout situations at the supplier and at customers.       
    
    for t in T:
        if t == 1:
            m.addConstr((Iit_1[1] + r[1]- sum(qikt[i,k,t] for i in N if i != 1 for k in K)) == Iit[1,t] )
        else:
            m.addConstr((Iit[1,t-1] + r[1]- sum(qikt[i,k,t] for i in N if i != 1 for k in K)) == Iit[1,t] )
    
    ##Constraint 3:-
    for i in N_dash:
        for t in T:
            if t == 1:
                m.addConstr(Iit_1[i] + sum(qikt[i,k,t] for k in K) - r[i] == Iit[i,t] )
            else:
                m.addConstr(Iit[i,t-1]+sum(qikt[i,k,t] for k in K) - r[i] == Iit[i,t])
                    
    ##Constraint 4:-
    for i in N:
        for t in T:
            m.addConstr( Iit[i,t] >= 0)
                    
    for i in N_dash:
        for t in T:
            m.addConstr(Iit[i,t]<=U[i])
            
    ######Constraint 5:-
    ######Constraints (5)–(7) ensure
    ######the OU policy requirements imposing that, if a customer is visited, the quantity delivered is such
    ######that the maximum inventory level is reached.
    for i in N_dash:
        for k in K:
            for t in T:
                if t == 1:
                    m.addConstr(sum(qikt[i,k,t] for k in K) >= U[i]*sum(zikt[i,k,t] for k in K)-Iit_1[i]) 
                else:
                    m.addConstr(sum(qikt[i,k,t] for k in K)>= U[i]*sum(zikt[i,k,t] for k in K)-Iit[i,t-1])
    ####Constraint 6:-        
    for i in N_dash:
        for t in T:
            if t == 1:
                m.addConstr(sum(qikt[i,k,t] for k in K)<= U[i] - Iit_1[i])
            else:
                m.addConstr(sum(qikt[i,k,t] for k in K)<= U[i] - Iit[i,t-1])
                 
    ##Constraint 7:-
    for i in N_dash:
        for k in K:
            for t in T:
                m.addConstr(qikt[i,k,t]<=U[i]*zikt[i,k,t])
    
    ####Constraint 8:-
    ####Constraints (8) are the vehicle capacity constraints.                    
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
    
    #####Constraints (9)–(11) are the routing constraints.
    ##Constraint 9:-
    ####Constraints (9) impose to visit each customer
    ####at most once in each time period, 
    for i in N_dash:
        for t in T:
            m.addConstr(sum(zikt[i,k,t] for k in K)<=1)
    
    ######Constraint 10:-
    ######constraints (10) are the degree constraints for each node and
    ######each vehicle in each time period,           
    for i in N:
        for k in K:
            for t in T:
                m.addConstr(sum(yijkt[i,j,k,t] for j in N  if j!=i)==zikt[i,k,t]) 
                m.addConstr(sum(yijkt[j,i,k,t] for j in N  if i!=j )==zikt[i,k,t])
    
    ##Constraint 13:-
    for i in N_dash:
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
                        
    #######Subtour Elimination Constraint:-
    ####Constraint 11:-
    #####(11) are the SECs for each vehicle route and each time
    #####period. Note that SECs (11) are stronger than those with right-hand side equal to |S| − 1. If we
    #####remove constraints (5) from model (k−A−ou), the resulting model (k−A−ml) applies for theML
    #####policy.                     
    for i in N_dash:
        for t in T:
            for k in K:
                for j in N:
                    if i!=j:
                        m.addConstr((Uikt[i,k,t]-Uikt[j,k,t] + Q*yijkt[i,j,k,t])<= Q - qikt[i,k,t])
    for i in N_dash:
        for t in T:
            for k in K:
                m.addConstr(Uikt[i,k,t]<=Q) and m.addConstr(Uikt[i,k,t]>=qikt[i,k,t])
    
    m.write('MTZ.lp')
    m.optimize()
    
    for v in m.getVars():
        if v.x > 0.01:
            print(v.varName, v.x)
    print('Objective:',round(m.objVal,2))
    
    return m.objVal

book = xlrd.open_workbook(os.path.join("DATA.xlsx"))
AA=[] 
sh = book.sheet_by_name("Sheet1")
i = 0
while True:
    try:
        sp = sh.cell_value(i,0)
        AA.append(sp)
        i = i + 1
    except IndexError:
        break   
Result={}
print(len(AA))
print(AA)
for i in AA:
    print(i)
    Result[i]=IRP(i)
workbook=xlsxwriter.Workbook('Result.xlsx')
worksheet=workbook.add_worksheet('Result')
i=0
for x in Result:
    f=AA[i]
    worksheet.write(i,0,f)
    e=Result[x]
    worksheet.write(i,1,e)    
    i+=1
workbook.close()
    

    

