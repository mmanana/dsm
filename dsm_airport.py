"""
Created on January 2017

@author:  Sergio Ortega & Mario Mañana 
"""

""" MAIN PROGRAM """

""" This file can be used as an example of optimization of airports energy demand based on DSM and Monte Carlo techniques. """
""" Application file...-."

""" INITILIZATION """

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

""" Number of Monte Carlo iterations """
#N =input("Number of Monte Carlo Simulations?: ")
#N=int(N)
N=1000

"""Minimum Quality of service """
#QoS =input("Minimum quality of service?: ")
#QoS=int(QoS)
QoS=25


"""PLOT THE ORIGINAL ELECTRIC LOAD PROFILE """
plt.figure(0)
dsm.PlotHourlyPower( LoadsList,'Original','Original')





"""INTRODUCTION """
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
print( 'Original indicators:')
print( 'Original Peak (kWh): {0:.2f}'.format(dsm.PeakValue( LoadsList)))
print( 'Original Load Factor: {0:.2f}'.format(dsm.LoadFactor(LoadsList)))
print( 'Original Ramping: {0:.2f}'.format(dsm.Ramping(LoadsList)))
print( 'Original Relation Peak To Valley: {0:.2f}'.format(dsm.PeakToValley(LoadsList)))
print( 'Original Energy Consumption Cost (€): {0:.2f}'.format(dsm.EnergyCost(LoadsList)))
print( '**************************************')
print( '**************************************')


"""SCENARIO 1"""
print( '')
print( '')
print( '**************************************')
print( '**************************************')
print( 'Scenario 1: Peak optimization')


"""LOAD SCHEDULING OPTIMIZATION"""
OptimizedList1 = dsm.PeakOptimization( LoadsList, N, QoS)


"""WRITE OPTIMIZED LOAD LIST IN EXCEL FILE"""
file = './file_peak.xlsx'
dsm.WriteExcel( file, OptimizedList1)       


"""CALCULATION OF OPTIMIZATED INDICATORS OF THE ELECTRIC LOAD PROFILE"""
print( '')
print( 'Optimizated indicators:')
print( 'Optimized Peak: {0:.2f}'.format(dsm.PeakValue(OptimizedList1)))
print( 'Load factor of scenario 1: {0:.2f}'.format( dsm.LoadFactor( OptimizedList1)))
print( 'Ramping of scenario 1: {0:.2f}'.format( dsm.Ramping( OptimizedList1)))
print( 'Relation Peak to valley of scenario 1: {0:.2f}'.format( dsm.PeakToValley( OptimizedList1)))
print( 'Energy consumption cost of scenario 1: {0:.2f}'.format(dsm.EnergyCost( OptimizedList1)))
print( '')


"""PLOT THE ORIGINAL AND OPTIMIZED ELECTRIC LOAD PROFILE"""
plt.figure(1)
dsm.PlotHourlyPower( LoadsList,'Original','Original')
dsm.PlotHourlyPower( OptimizedList1,'Peak optimization','Optimized')


"""CALCULATION OF QUALITY SERVICES"""
print( 'Quality of service calculation:')
QoSTotal=dsm.QoSTotal( LoadsList,OptimizedList1)

print( '**************************************')
print( '**************************************')





"""SCENARIO 2"""
print( '')
print( '')
print( '**************************************')
print( '**************************************')
print( 'Scenario 2: Energy consumption cost optimization')


"""LOAD SCHEDULING OPTIMIZATION"""
OptimizedList2 = dsm.CostOptimization( LoadsList, N, QoS)


"""WRITE OPTIMIZED LOAD LIST IN EXCEL FILE"""
file = './file_cost.xlsx'
dsm.WriteExcel( file, OptimizedList2)       


"""CALCULATION OF OPTIMIZATED INDICATORS OF THE ELECTRIC LOAD PROFILE"""
print( '')
print( 'Optimizated indicators:')
print( 'Optimized Peak of scenario 2: {0:.2f}'.format(dsm.PeakValue(OptimizedList2)))
print( 'Load factor of scenario 2: {0:.2f}'.format( dsm.LoadFactor( OptimizedList2)))
print( 'Ramping of scenario 2: {0:.2f}'.format( dsm.Ramping( OptimizedList2)))
print( 'Relation Peak to valley of scenario 2: {0:.2f}'.format( dsm.PeakToValley( OptimizedList2)))
print( 'Energy consumption cost of scenario 2: {0:.2f}'.format(dsm.EnergyCost( OptimizedList2)))
print( '')


"""PLOT THE ORIGINAL AND OPTIMIZED ELECTRIC LOAD PROFILE"""
plt.figure(2)
dsm.PlotHourlyPower( LoadsList,'Original','Original')
dsm.PlotHourlyPower( OptimizedList2,'Cost optimization','Optimized')


"""CALCULATION OF QUALITY SERVICES"""
print( 'Quality of service calculation:')
QoSTotal=dsm.QoSTotal( LoadsList,OptimizedList2)

print( '**************************************')
print( '**************************************')


#
#
"""SCENARIO 3"""
print( '')
print( '')
print( '**************************************')
print( '**************************************')
print( 'Scenario 3:Ramping optimization')


"""LOAD SCHEDULING OPTIMIZATION"""
OptimizedList3 = dsm.RampingOptimization( LoadsList, N, QoS)


"""WRITE OPTIMIZED LOAD LIST IN EXCEL FILE"""
file = './file_ramping.xlsx'
dsm.WriteExcel( file, OptimizedList3)       


"""CALCULATION OF OPTIMIZATED INDICATORS OF THE ELECTRIC LOAD PROFILE"""
print( '')
print( 'Optimizated indicators:')
print( 'Optimized Peak of scenario 3: {0:.2f}'.format(dsm.PeakValue(OptimizedList3)))
print( 'Load factor of scenario 3: {0:.2f}'.format( dsm.LoadFactor( OptimizedList3)))
print( 'Ramping of scenario 3: {0:.2f}'.format( dsm.Ramping( OptimizedList3)))
print( 'Relation Peak to valley of scenario 3: {0:.2f}'.format( dsm.PeakToValley( OptimizedList3)))
print( 'Energy consumption cost of scenario 3: {0:.2f}'.format(dsm.EnergyCost( OptimizedList3)))
print( '')


"""PLOT THE ORIGINAL AND OPTIMIZED ELECTRIC LOAD PROFILE"""
plt.figure(3)
dsm.PlotHourlyPower( LoadsList,'Original','Original')
dsm.PlotHourlyPower( OptimizedList3,'Ramping optimization','Optimized')


"""CALCULATION OF QUALITY SERVICES"""
print( 'Quality of service calculation:')
QoSTotal=dsm.QoSTotal( LoadsList,OptimizedList3)

print( '**************************************')
print( '**************************************')

#


"""SCENARIO 4"""
print( '')
print( '')
print( '**************************************')
print( '**************************************')
print( 'Scenario 4: Peak to Valley optimization')


"""LOAD SCHEDULING OPTIMIZATION"""
OptimizedList4 = dsm.OptimizacionPeaktoValley( LoadsList, N, QoS)


"""WRITE OPTIMIZED LOAD LIST IN EXCEL FILE"""
file = './file_peaktovalley.xlsx'
dsm.WriteExcel( file, OptimizedList4)       


"""CALCULATION OF OPTIMIZATED INDICATORS OF THE ELECTRIC LOAD PROFILE"""
print( '')
print( 'Optimizated indicators:')
print( 'Optimized Peak of scenario 4: {0:.2f}'.format(dsm.PeakValue(OptimizedList4)))
print( 'Load factor of scenario 4: {0:.2f}'.format( dsm.LoadFactor( OptimizedList4)))
print( 'Ramping of scenario 4: {0:.2f}'.format( dsm.Ramping( OptimizedList4)))
print( 'Relation Peak to valley of scenario 4: {0:.2f}'.format( dsm.PeakToValley( OptimizedList4)))
print( 'Energy consumption cost of scenario 4: {0:.2f}'.format(dsm.EnergyCost( OptimizedList4)))
print( '')


"""PLOT THE ORIGINAL AND OPTIMIZED ELECTRIC LOAD PROFILE"""
plt.figure(4)
dsm.PlotHourlyPower( LoadsList,'Original','Original')
dsm.PlotHourlyPower( OptimizedList4,'Peak-to-Valley optimization','Optimized')


"""CALCULATION OF QUALITY SERVICES"""
print( 'Quality of service calculation:')
QoSTotal=dsm.QoSTotal( LoadsList,OptimizedList4)

print( '**************************************')
print( '**************************************')



"""SCENARIO 5"""
print( '')
print( '')
print( '**************************************')
print( '**************************************')
print( 'Scenario 5: Load factor optimization')


"""LOAD SCHEDULING OPTIMIZATION"""
OptimizedList5 = dsm.OptimizacionLoadFactor( LoadsList, N, QoS)


"""WRITE OPTIMIZED LOAD LIST IN EXCEL FILE"""
file = './file_loadfactor.xlsx'
dsm.WriteExcel( file, OptimizedList5)       


"""CALCULATION OF OPTIMIZATED INDICATORS OF THE ELECTRIC LOAD PROFILE"""
print( '')
print( 'Optimizated indicators:')
print( 'Optimized Peak of scenario 5: {0:.2f}'.format(dsm.PeakValue(OptimizedList5)))
print( 'Load factor of scenario 5: {0:.2f}'.format( dsm.LoadFactor( OptimizedList5)))
print( 'Ramping of scenario 5: {0:.2f}'.format( dsm.Ramping( OptimizedList5)))
print( 'Relation Peak to valley of scenario 5: {0:.2f}'.format( dsm.PeakToValley( OptimizedList5)))
print( 'Energy consumption cost of scenario 5: {0:.2f}'.format(dsm.EnergyCost( OptimizedList5)))
print( '')


"""PLOT THE ORIGINAL AND OPTIMIZED ELECTRIC LOAD PROFILE"""
plt.figure(5)
dsm.PlotHourlyPower( LoadsList,'Original','Original')
dsm.PlotHourlyPower( OptimizedList5,'Load factor optimization','Optimized')


"""CALCULATION OF QUALITY SERVICES"""
print( 'Quality of service calculation:')
QoSTotal=dsm.QoSTotal( LoadsList,OptimizedList5)

print( '**************************************')
print( '**************************************')







