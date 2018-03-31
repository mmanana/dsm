"""
Created on December 2017

@author:  Sergio Ortega & Mario Mañana 
"""

""" MAIN PROGRAM"""



""" INITIALIZATION """

import dsm_loads as dsm
import numpy as np
import xlsxwriter
import copy
import matplotlib.pyplot as plt



""" INPUT DATA """

""" Function LoadData """
# Function to load loads list from excel file 
file = 'loads-airport.xlsx'
LoadsList = dsm.LoadData( file) 

""" Number of Monte Carlo iterations"""
#N =input("Number of Monte Carlo Simulations?: ")
#N=int(N)
N=35000

"""Minimum Quality of service"""
#QoS =input("Minimum quality of service?: ")
#QoS=int(QoS)
QoS=60


#"""PLOT THE ORIGINAL ELECTRIC LOAD PROFILE"""
#plt.figure(0)
#dsm.PlotHourlyPower( LoadsList,'Original','Original')





"""INTRODUCTION"""
print( '')
print( '')
print( '**************************************')
print( 'AIRPORT LOAD SCHEDULING BASED ON DSM ')
print( 'STRATEGIES AND  MONTE CARLO METHODOLOGY')
print( '**************************************')
print( '**************************************')
print( 'Simulation day: 18/02/2015')
print( 'Number of Monte Carlo simulations: {0:.0f}'.format(N))
print( 'Minimum Quality of service: {0:.0f}'.format(QoS))
print( '**************************************')
print( '**************************************')


"""ORIGINAL INDICATORS"""
print( '')
print( '')
print( '**************************************')
print( '**************************************')
print( 'Original Energy Consumption Cost (€): {0:.2f}'.format(dsm.EnergyCost( LoadsList)))
print( '**************************************')
print( '**************************************')





"""LOAD SCHEDULING OPTIMIZATION"""
OptimizedList = dsm.CostOptimization( LoadsList, N, QoS)


"""WRITE OPTIMIZED LOAD LIST IN EXCEL FILE"""
file = './file_optimized-loads.xlsx'
dsm.WriteExcel( file, OptimizedList)       


"""CALCULATION OF OPTIMIZATED INDICATORS OF THE ELECTRIC LOAD PROFILE"""
print( '')
print( '**************************************')
print( '**************************************')
print( 'Optimized Energy Consumption Cost: {0:.2f}'.format(dsm.EnergyCost( OptimizedList)))
print( '**************************************')
print( '**************************************')
print( '')


"""CALCULATION OF QUALITY SERVICES"""
print( '**************************************')
print( '**************************************')
print( 'Quality of service calculation:')
QoSTotal=dsm.QoSTotal( LoadsList,OptimizedList)
print( '**************************************')
print( '**************************************')


"""PLOT THE ORIGINAL AND OPTIMIZED ELECTRIC LOAD PROFILE"""
plt.figure(1)
dsm.PlotHourlyPower( LoadsList,'Cost Optimization','Original')
dsm.PlotHourlyPower( OptimizedList,'Cost Optimization','Optimized')






