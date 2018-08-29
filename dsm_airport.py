"""
Created on December 2017
Last Update: August 2018

@author:  Sergio Ortega & Mario Mañana 
"""

""" MAIN PROGRAM"""



""" INITIALIZATION """

import dsm_loads as dsm
import numpy as np
import xlsxwriter
import copy
import sys
import matplotlib.pyplot as plt



""" INPUT DATA """

""" Function LoadData """
# Function to load loads list from excel file 
file = 'loads-airport.xlsx'
LoadsList = dsm.LoadData( file) 

""" Number of iterations"""
#N =input("Number of main iterations?: ")
#N=int(N)
N=100

"""Minimum Quality of service"""
#QoS =input("Minimum quality of service?: ")
#QoS=int(QoS)
QoS=60


#"""PLOT THE ORIGINAL ELECTRIC LOAD PROFILE"""
#plt.figure(0)
#dsm.PlotHourlyPower( LoadsList,'Original','Original')
"""Types of optimization """
#             0     1
TypesOfOpt=['MC', 'SA']
TOO=TypesOfOpt[1]


"""INTRODUCTION"""
print( '')
print( '')
print( '*********************************************************************')
print( 'AIRPORT LOAD SCHEDULING BASED ON DSM ')
print( 'STRATEGIES AND  OPTIMIZATION METHODOLOGIES')
print( '*********************************************************************')
print( 'Python: ' + str(sys.version))
print( '*********************************************************************')
print( 'Type of Optimization: ' + TOO)
print( 'Simulation day: 18/02/2015')
print( 'Number of simulations: {0:.0f}'.format(N))
print( 'Minimum Quality of service: {0:.0f}'.format(QoS))
print( '*********************************************************************')
print( '*********************************************************************')


"""ORIGINAL INDICATORS"""
print( '')
print( '')
print( '*********************************************************************')
print( '*********************************************************************')
print( 'Original Energy Consumption Cost (€): {0:.2f}'.format(dsm.EnergyCost( LoadsList)))
print( '*********************************************************************')
print( '*********************************************************************')



"""LOAD SCHEDULING OPTIMIZATION"""
if TOO == 'MC':
    print(' Monte Carlo Optimization ')
    OptimizedList = dsm.CostOptimizationMC( LoadsList, N, QoS)
elif TOO == 'SA':
    print(' Simulated Annealing ')
    OptimizedList = dsm.CostOptimizationSA( LoadsList, N, QoS)
else:
    print('Who knows...')
    

"""WRITE OPTIMIZED LOAD LIST IN EXCEL FILE"""
file = './file_optimized-loads.xlsx'
dsm.WriteExcel( file, OptimizedList)       


"""CALCULATION OF OPTIMIZATED INDICATORS OF THE ELECTRIC LOAD PROFILE"""
print( '')
print( '*********************************************************************')
print( '*********************************************************************')
print( 'Optimized Energy Consumption Cost: {0:.2f}'.format(dsm.EnergyCost( OptimizedList)))
print( '*********************************************************************')
print( '*********************************************************************')
print( '')


"""CALCULATION OF QUALITY SERVICES"""
print( '*********************************************************************')
print( '*********************************************************************')
print( 'Quality of service calculation:')
QoSTotal=dsm.QoSTotal( LoadsList,OptimizedList)
print( '*********************************************************************')
print( '*********************************************************************')


"""PLOT THE ORIGINAL AND OPTIMIZED ELECTRIC LOAD PROFILE"""
plt.figure(1)
dsm.PlotHourlyPower( LoadsList,'Cost Optimization','Original')
dsm.PlotHourlyPower( OptimizedList,'Cost Optimization','Optimized')






