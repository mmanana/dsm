"""
Created on December 2017

@author:  Sergio Ortega & Mario Mañana 
"""

"""*********************************************************** """

""" LIBRARIES"""
import time
import numpy as np
import pandas as pd
import random 
import copy
import matplotlib.pyplot as plt
import xlsxwriter


""" TIME SLOTS """
hour_list = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95]


"""*********************************************************** """

""" CLASSES DEFINITION"""

class BaseLoad:
    sName = ' '
    sBuilding = ' '
    sSubzone= ' '
    sFacility= ' '
    sReference = ' '
    fPowerQh = 0.0
    fEnergyQh = 0.0
    # Load status: RUN:1 STOP:0
    bStatus = 0
    fPN = 0.0
    fEtot = 0.0
    sType = ' '
    
    def __init__(self, name):
        self.sName = name 
        self.fPowerQh = np.array(np.zeros([1,96]))
        self.fEnergyQh = np.array(np.zeros([1,96]))
        self.bStatus = np.array(np.zeros([1,96]))
        
    def SetReference( self, ref):
        self.sReference = ref
        
    def SetType( self, ref):
        self.sType = ref
        
    def SetBuilding( self, ref):
        self.sBuilding = ref
        
    def SetSubzone( self, ref):
        self.sSubzone = ref
        
    def SetFacility( self, ref):
        self.sFacility = ref
        
    def SetPowerQh( self, data):
        self.fa = data[0]
        for i in range(0,96):
           self.fPowerQh.itemset(i,data[i])
           
    def SetStatus( self, Status):
        self.bStatus = Status       
        
    def SetPN( self, PN):
        self.fPN = PN
        
    def SetEtot( self, Etot):
        self.fEtot = Etot
        

class InterruptibleLoad(BaseLoad):
    clasename = 'InterruptibleLoad'    
        
        
"""STATIC LOAD"""
class StaticLoad (BaseLoad):
    clasename = 'StaticLoad'
    
    
class ElasticLoad (BaseLoad):
    clasename = 'ElasticLoad'    

    # SubType
    sSubType = ' '
    # Minimum scale
    fScaleMin = 1.0
    # scale
    fScale = 1.0

    def SetSubType( self, SubType):
        self.sSubType = SubType
        
    def SetScaleMin( self, ScaleMin):
        self.fScaleMin = ScaleMin
        
    def SetScale( self, Scale):
        self.fScale = Scale
  

class ShiftableLoad(BaseLoad):
    clasename = 'ShiftableLoad'        
    
    Qhstart = 0
    Qhend = 0
    Qhstartmin = 0
    Qhstartmax = 0

    def SetQhInterval( self, Qhs, Qhe, Qhsmin, Qhsmax):
        self.Qhstart = Qhs
        self.Qhend = Qhe
        self.Qhstartmin = Qhsmin
        self.Qhstartmax = Qhsmax
    
"""*********************************************************** """
    
    
  
"""*********************************************************** """   

""" FUNCTION TO LOAD LOADSLIST AND CREATE TABLE MODEL"""

def LoadData(file):
    
    tableStatic= pd.read_excel (file,
                               sheetname='STATIC',
                               header= 0,
                               convert_float= True)
    Nst=len(tableStatic.index)
    
    
    tableShiftabled= pd.read_excel (file,
                               sheetname='SHIFTABLED',
                               header= 0,
                               convert_float= True)
    Nsh=len(tableShiftabled.index)
    
    
    tableElastic= pd.read_excel (file,
                               sheetname='ELASTIC',
                               header= 0,
                               convert_float= True)
    Ne=len(tableElastic.index)
    
    
    tableInterruptible= pd.read_excel (file,
                               sheetname='INTERRUPTIBLE',
                               header= 0,
                               convert_float= True)
    Ni=len(tableInterruptible.index)
    

    LoadsList= []
  
    # Counter
    h=0
    
    # Static Loads
    for i in range(0,Nst):
        sName = tableStatic[:]['Name'].iloc[i]
        LoadsList.append(StaticLoad(sName))
        sReference = tableStatic[:]['Reference'].iloc[i]
        LoadsList[h].SetReference( sReference)
        sBuilding = tableStatic[:]['Building'].iloc[i]
        LoadsList[h].SetBuilding( sBuilding)
        sSubzone = tableStatic[:]['Subzone'].iloc[i]
        LoadsList[h].SetSubzone( sSubzone)
        sFacility = tableStatic[:]['Facility'].iloc[i]
        LoadsList[h].SetFacility( sFacility)
        Status = tableStatic[:]['Status'].iloc[i]
        LoadsList[h].SetStatus( Status)
        fPN = tableStatic[:]['PN'].iloc[i]
        LoadsList[h].SetPN( fPN)
        fEtot = tableStatic[:]['Etot'].iloc[i]
        LoadsList[h].SetEtot( fEtot)        
        LoadsList[h].SetType( 'Static')
        data=np.array(tableStatic[:][hour_list].iloc[i])
        LoadsList[h].SetPowerQh( data)
        h = h+1

    #  Shiftable Loads 
    for i in range(0,Nsh):
        sName = tableShiftabled[:]['Name'].iloc[i]
        LoadsList.append( ShiftableLoad(sName))
        sReference = tableShiftabled[:]['Reference'].iloc[i]
        LoadsList[h].SetReference( sReference)
        sBuilding = tableShiftabled[:]['Building'].iloc[i]
        LoadsList[h].SetBuilding( sBuilding)
        sSubzone = tableShiftabled[:]['Subzone'].iloc[i]
        LoadsList[h].SetSubzone( sSubzone)
        sFacility = tableShiftabled[:]['Facility'].iloc[i]
        LoadsList[h].SetFacility( sFacility)
        Status = tableShiftabled[:]['Status'].iloc[i]
        LoadsList[h].SetStatus( Status)
        fPN = tableShiftabled[:]['PN'].iloc[i]
        LoadsList[h].SetPN( fPN)
        fEtot = tableShiftabled[:]['Etot'].iloc[i]
        LoadsList[h].SetEtot( fEtot)        
        Qhs = tableShiftabled[:]['Qhstart'].iloc[i]
        Qhe = tableShiftabled[:]['Qhend'].iloc[i]
        Qhsmin = tableShiftabled[:]['Qhstartmin'].iloc[i]
        Qhsmax = tableShiftabled[:]['Qhstartmax'].iloc[i]
        LoadsList[h].SetQhInterval( Qhs, Qhe, Qhsmin, Qhsmax)
        LoadsList[h].SetType( 'Shiftable')
        data=np.array(tableShiftabled[:][hour_list].iloc[i])
        LoadsList[h].SetPowerQh( data) 
        h = h+1
    
    # Elastic Loads
    for i in range(0,Ne):
        sName = tableElastic[:]['Name'].iloc[i]
        LoadsList.append( ElasticLoad(sName))
        sReference = tableElastic[:]['Reference'].iloc[i]
        LoadsList[h].SetReference( sReference)
        sBuilding = tableElastic[:]['Building'].iloc[i]
        LoadsList[h].SetBuilding( sBuilding)
        sSubzone = tableElastic[:]['Subzone'].iloc[i]
        LoadsList[h].SetSubzone( sSubzone)
        sFacility = tableElastic[:]['Facility'].iloc[i]
        LoadsList[h].SetFacility( sFacility)
        sSubType = tableElastic[:]['SubType'].iloc[i]
        LoadsList[h].SetSubType( sSubType)
        Scale = tableElastic[:]['Scale'].iloc[i]
        LoadsList[h].SetScale( Scale)        
        ScaleMin = tableElastic[:]['ScaleMin'].iloc[i]
        LoadsList[h].SetScaleMin( ScaleMin)        
        Status = tableElastic[:]['Status'].iloc[i]
        LoadsList[h].SetStatus( Status)
        fPN = tableElastic[:]['PN'].iloc[i]
        LoadsList[h].SetPN( fPN)
        fEtot = tableElastic[:]['Etot'].iloc[i]
        LoadsList[h].SetEtot( fEtot)
        LoadsList[h].SetType( 'Elastic')
        data=np.array(tableElastic[:][hour_list].iloc[i])
        LoadsList[h].SetPowerQh( data)  
        h = h+1

        
        # Interruptible Loads
    for i in range(0,Ni):
        sName = tableInterruptible[:]['Name'].iloc[i]
        LoadsList.append( InterruptibleLoad(sName))
        sReference = tableInterruptible[:]['Reference'].iloc[i]
        LoadsList[h].SetReference( sReference)
        sBuilding = tableInterruptible[:]['Building'].iloc[i]
        LoadsList[h].SetBuilding( sBuilding)
        sSubzone = tableInterruptible[:]['Subzone'].iloc[i]
        LoadsList[h].SetSubzone( sSubzone)
        sFacility = tableInterruptible[:]['Facility'].iloc[i]
        LoadsList[h].SetFacility( sFacility)
        Status = tableInterruptible[:]['Status'].iloc[i]
        LoadsList[h].SetStatus( Status)
        fPN = tableInterruptible[:]['PN'].iloc[i]
        LoadsList[h].SetPN( fPN)
        fEtot = tableInterruptible[:]['Etot'].iloc[i]
        LoadsList[h].SetEtot( fEtot)
        LoadsList[h].SetType( 'Interruptible')
        data=np.array(tableInterruptible[:][hour_list].iloc[i])
        LoadsList[h].SetPowerQh( data)  
        h = h+1

    return LoadsList
    
"""*********************************************************** """


"""*********************************************************** """

""" FUNCTION TO CALCULATE AGGREGATED POWER OF LOADS LIST IN EACH TIME SLOT """

def AggregatePower( LoadsList):    
    fAggregatedPower = np.array(np.zeros([1,96]))
    for Load in LoadsList:
        fAggregatedPower = fAggregatedPower + Load.fPowerQh
    return fAggregatedPower

"""*********************************************************** """



"""*********************************************************** """

""" FUNCTION TO PLOT AGGREATED POWER OF LOADS LIST """

def PlotHourlyPower( Loadslist,title1,title2):
    N = 96 
    x = np.arange(N)
    fAggregatedPower=AggregatePower(Loadslist)
    y = tuple( fAggregatedPower.reshape(1,-1)[0])
    plt.plot(x, y,label=title2)
    plt.legend(loc="upper right")
    plt.xlabel("Slot")
    plt.ylabel("Power demand")
    plt.title(title1)
    return

"""*********************************************************** """




"""*********************************************************** """ 

""" ALGORITHM MONTE CARLO """

def MonteCarlo( LoadsList):    
    LoadsList1 = copy.deepcopy( LoadsList)
    for i in range(len(LoadsList1)):
        if LoadsList1[i].sType == 'Interruptible':
            r = random.randint(1,100)
            if r < 50:            
                LoadsList1[i].fPowerQh = np.array(np.zeros([1,96]))
                LoadsList1[i].SetStatus( 0)
        elif LoadsList1[i].sType == 'Shiftable':
            h = random.randint( LoadsList1[i].Qhstartmin, LoadsList1[i].Qhstartmax)
            D = LoadsList1[i].Qhend - LoadsList1[i].Qhstart 
            LoadsList1[i].fPowerQh = np.array(np.zeros([1,96]))
            for k in range( h, h+D+1):
                LoadsList1[i].fPowerQh[0,k] = LoadsList[i].fPowerQh[0,k-h+LoadsList1[i].Qhstart]
                LoadsList1[i].Qhstart=h
                LoadsList1[i].Qhend=h+D 
        elif LoadsList1[i].sType == 'Elastic':
            if LoadsList1[i].sSubType == 'modulable':
                fs=LoadsList1[i].fScale
                fsmin = LoadsList1[i].fScaleMin
                fs = random.uniform( fsmin, fs)
                LoadsList1[i].SetScale( fs)
                for h in range( 0, 96):
                    LoadsList1[i].fPowerQh[0,h] = fs*LoadsList[i].fPowerQh[0,h]
    return LoadsList1
"""*********************************************************** """     





"""*********************************************************** """
 

""" FUNCTION: ENERGY COST CONSUMPTION OPTIMIZATION"""
  
def EnergyCost( LoadsList):
    
    faP = AggregatePower( LoadsList)
    
    #ENERGY COST BY PERIODS (EUROS/kWH)
    P6=0.049571
    P2=0.087605
    P1=0.110585
    
    #PERIOD P0 TO P31
    suma=0
    energy=0
   
    EcostP0P31=0
    for i in range(0,32):
        suma = suma+faP[0,i]
        i = i + 1
    energy=suma/4
    EcostP0P31=energy*P6
        
    #PERIOD P32 TO P39
    suma=0
    energy=0
    EcostP32P39=0
    for i in range(32,40):
        suma = suma+faP[0,i]
        i = i + 1
    energy=suma/4
    EcostP32P39=energy*P2
        
    #PERIOD P40 TO P51
    suma=0
    energy=0
    EcostP40P51=0
    for i in range(40,52):
        suma = suma+faP[0,i]
        i = i + 1
    energy=suma/4
    EcostP40P51=energy*P1
        
    #PERIOD P52 TO P71
    suma=0
    energy=0
    EcostP52P71=0
    for i in range(52,72):
        suma = suma+faP[0,i]
        i = i + 1
    energy=suma/4
    EcostP52P71=energy*P2
        
    #PERIOD P72 TO P83
    suma=0
    energy=0
    EcostP72P83=0
    for i in range(72,84):
        suma = suma+faP[0,i]
        i = i + 1
    energy=suma/4
    EcostP72P83=energy*P1
        
    #PERIOD P72 TO P83
    suma=0
    energy=0
    EcostP84P95=0
    for i in range(84,96):
        suma = suma+faP[0,i]
        i = i + 1
    energy=suma/4
    EcostP84P95=energy*P2
        
    Ecosttotal=EcostP0P31+EcostP32P39+EcostP40P51+EcostP52P71+EcostP72P83+EcostP84P95
    
    return Ecosttotal

"""*********************************************************** """ 


    
"""*********************************************************** """ 

def CostOptimization( LoadsList, Niterations,QoS):

    Ecostmin = np.empty([1,Niterations+1])
    Ecosttotal = np.empty([1,Niterations+1])

    h=0
    Ecostmin[0,h] = EnergyCost( LoadsList) 
    Ecosttotal[0,h] = EnergyCost( LoadsList)

    LoadsList1 = copy.deepcopy( LoadsList)

    for h in range(0,Niterations):
        Ecost1 = EnergyCost( LoadsList1)
        LoadsList2 = copy.deepcopy(MonteCarlo( LoadsList))
        Ecost2 = EnergyCost( LoadsList2)
        Ecosttotal[0,h+1]=Ecost2
        QoSLoadlist2=QoSTotalB (LoadsList,LoadsList2)
        if (QoSLoadlist2 > QoS) & (Ecost2 < Ecost1):
            Ecostmin[0,h+1]=Ecost2
            LoadsList1 = copy.deepcopy( LoadsList2)
        else:
            Ecostmin[0,h+1]=Ecost1
            LoadsList1 = copy.deepcopy( LoadsList1)
        if ((h==9) or (h==24) or (h==49) or (h==99) or (h==249) or (h==499)):
                print (h+1)
                print(Ecostmin[0,h+1])
        if  ((h==999) or (h==2499) or (h==4999) or (h==9999) ):
                print (h+1)
                print(Ecostmin[0,h+1])
        if  ((h==14999) or (h==19999) or (h==24999)):
                print (h+1)
                print(Ecostmin[0,h+1])
        if  ((h==29999) or (h==34999) or (h==39999)):
                print (h+1)
                print(Ecostmin[0,h+1])
        h=h+1
    return LoadsList1   
  
"""*********************************************************** """ 

      




          

"""*********************************************************** """  

""" FUNCTION TO CALCULATE QUALITY OF SERVICE OF INTERRUPTIBLE LOADS """     
        
def QoSInterruptible ( LoadsList):
    InterruptibleTotal =0
    InterruptibleOn=0
    
    for i in range(len(LoadsList)):
        if LoadsList[i].sType == 'Interruptible':    
            if LoadsList[i].bStatus == 1:
                    InterruptibleOn=InterruptibleOn+1
                    InterruptibleTotal=InterruptibleTotal+1    
            else:
                    InterruptibleTotal=InterruptibleTotal+1
    QoSInterrupt=100*(InterruptibleOn/InterruptibleTotal)
    return QoSInterrupt  


    
"""*********************************************************** """  
""" FUNCTION TO CALCULATE QUALITY OF SERVICE OF ELASTIC LOADS """    
         
def QoSElastic( LoadsList):
    fsmin=0.8
    FactorElast=0
    Count=0
    for i in range(len(LoadsList)):
            if LoadsList[i].sType == 'Elastic':
                if LoadsList[i].sSubType== 'modulable':
                            FactorElast=FactorElast+LoadsList[i].fScale
                            Count=Count+1
    FactorElast=FactorElast/Count
    QoSElast=100*(1-((1-FactorElast)/(1-fsmin)))
    return QoSElast

"""*********************************************************** """  


""" FUNCTION TO CALCULATE QUALITY OF SERVICE OF SHIFTABLE LOADS  """    
         
def QoSShiftable( LoadsList,OptimizedList):
 
    Count=0
    QoSShiftload=0
    QoSShift=0
    for i in range(len(OptimizedList)):
            if OptimizedList[i].sType == 'Shiftable':
                QoSShiftload=100*(1-((abs(LoadsList[i].Qhstart-OptimizedList[i].Qhstart))/(LoadsList[i].Qhstartmax-LoadsList[i].Qhstartmin)))
                QoSShift=QoSShift+QoSShiftload  
                Count=Count+1
    QoSShift=QoSShift/Count
    return QoSShift

"""*********************************************************** """  

""" FUNCTION TO CALCULATE QUALITY OF SERVICE OF TOTAL LOADS """    
         
def QoSTotal( LoadsList, OptimizedList):
   
    QoSInterrupt=QoSInterruptible ( OptimizedList)
    QoSElast=QoSElastic( OptimizedList)  
    QoSShift=QoSShiftable( LoadsList, OptimizedList)
    
    QoSTot=(QoSElast+QoSInterrupt+QoSShift)/3
    
    print('Quality of service of interruptible loads is (%):',QoSInterrupt)
    print('Quality of service of elastic loads is (%):',QoSElast)
    print('Quality of service of shiftable loads is (%):',QoSShift)
    print('Total quality of service of loads is (%):',QoSTot)   
    return QoSTot

    
"""*********************************************************** """  

""" FUNCTION TO CALCULATE QUALITY OF SERVICE OF TOTAL LOADS WITHOUT PRINT RESULTS"""    
         
def QoSTotalB( LoadsList, OptimizedList):
   
    QoSInterrupt=QoSInterruptible ( OptimizedList)
    QoSElast=QoSElastic( OptimizedList)  
    QoSShift=QoSShiftable( LoadsList, OptimizedList)
    
    QoSTot=(QoSElast+QoSInterrupt+QoSShift)/3
    
    return QoSTot

    
"""*********************************************************** """  









"""*********************************************************** """ 

""" FUNCTION TO WRITE OPTIMIZED LOAD LIST IN EXCEL FILE """

def WriteExcel( file, LoadsList):

    wb = xlsxwriter.Workbook( file)
    bold = wb.add_format({'bold': 1})
    ws_static = wb.add_worksheet('STATIC')
    ws_static.write('A1', 'Reference', bold)
    ws_static.write('B1', 'Name', bold)
    ws_static.write('C1', 'Building', bold) 
    ws_static.write('D1', 'Subzone', bold)
    ws_static.write('E1', 'Facility', bold)
    ws_static.write('F1', 'Status', bold)
    ws_static.write('G1', 'PN', bold)
    ws_static.write('H1', 'Etot', bold)
    for h in range( 0, 96):
        ws_static.write_string( 0, 8+h, str(h), bold) 

    ws_interruptible = wb.add_worksheet('INTERRUPTIBLE')
    ws_interruptible.write('A1', 'Reference', bold)
    ws_interruptible.write('B1', 'Name', bold)
    ws_interruptible.write('C1', 'Building', bold)
    ws_interruptible.write('D1', 'Subzone', bold)
    ws_interruptible.write('E1', 'Facility', bold)
    ws_interruptible.write('F1', 'Status', bold)
    ws_interruptible.write('G1', 'PN', bold)
    ws_interruptible.write('H1', 'Etot', bold)
    for h in range( 0, 96):
        ws_interruptible.write_string( 0, 8+h, str(h), bold) 

    ws_shiftable = wb.add_worksheet('SHIFTABLE')
    ws_shiftable.write('A1', 'Reference', bold)
    ws_shiftable.write('B1', 'Name', bold)
    ws_shiftable.write('C1', 'Building', bold)
    ws_shiftable.write('D1', 'Subzone', bold)
    ws_shiftable.write('E1', 'Facility', bold)
    ws_shiftable.write('F1', 'Status', bold)    
    ws_shiftable.write('G1', 'PN', bold)
    ws_shiftable.write('H1', 'Etot', bold)
    ws_shiftable.write('I1', 'Qhstart', bold)
    ws_shiftable.write('J1', 'Qhend', bold)
    ws_shiftable.write('K1', 'Qhstartmin', bold)
    ws_shiftable.write('L1', 'Qhstartmax', bold)
    for h in range( 0, 96):
        ws_shiftable.write_string( 0, 12+h, str(h), bold) 

    ws_elastic = wb.add_worksheet('ELASTIC')
    ws_elastic.write('A1', 'Reference', bold)
    ws_elastic.write('B1', 'Name', bold)
    ws_elastic.write('C1', 'Building', bold)    
    ws_elastic.write('D1', 'Subzone', bold)
    ws_elastic.write('E1', 'Facility', bold)
    ws_elastic.write('F1', 'SubType', bold)
    ws_elastic.write('G1', 'Scale', bold)
    ws_elastic.write('H1', 'ScaleMin', bold)
    ws_elastic.write('I1', 'Status', bold)
    ws_elastic.write('J1', 'PN', bold)
    ws_elastic.write('K1', 'Etot', bold)
    for h in range( 0, 96):
        ws_elastic.write_string( 0, 11+h, str(h), bold) 

    row_static = 0
    row_interruptible = 0
    row_shiftable = 0
    row_elastic = 0
    for Load in LoadsList:
        if Load.sType == 'Static':
            row_static = row_static+1
            ws_static.write_string(row_static, 0, Load.sReference)        
            ws_static.write_string(row_static, 1, Load.sName)
            ws_static.write_string(row_static, 2, Load.sBuilding)
            ws_static.write_string(row_static, 3, Load.sSubzone)
            ws_static.write_string(row_static, 4, Load.sFacility)
            ws_static.write_number(row_static, 5, Load.bStatus)
            ws_static.write_number(row_static, 6, Load.fPN)
            ws_static.write_number(row_static, 7, Load.fEtot)
            for h in range( 0, 96):
                ws_static.write_number(row_static, 8+h, Load.fPowerQh[0,h])    
        elif Load.sType == 'Interruptible':
            row_interruptible = row_interruptible+1
            ws_interruptible.write_string(row_interruptible, 0, Load.sReference)        
            ws_interruptible.write_string(row_interruptible, 1, Load.sName)
            ws_interruptible.write_string(row_interruptible, 2, Load.sBuilding)
            ws_interruptible.write_string(row_interruptible, 3, Load.sSubzone)
            ws_interruptible.write_string(row_interruptible, 4, Load.sFacility)
            ws_interruptible.write_number(row_interruptible, 5, Load.bStatus)
            ws_interruptible.write_number(row_interruptible, 6, Load.fPN)
            ws_interruptible.write_number(row_interruptible, 7, Load.fEtot)
            for h in range( 0, 96):
                ws_interruptible.write_number(row_interruptible, 8+h, Load.fPowerQh[0,h])  
        elif Load.sType == 'Shiftable':
            row_shiftable = row_shiftable + 1
            ws_shiftable.write_string(row_shiftable, 0, Load.sReference)        
            ws_shiftable.write_string(row_shiftable, 1, Load.sName)
            ws_shiftable.write_string(row_shiftable, 2, Load.sBuilding)
            ws_shiftable.write_string(row_shiftable, 3, Load.sSubzone)
            ws_shiftable.write_string(row_shiftable, 4, Load.sFacility)
            ws_shiftable.write_number(row_shiftable, 5, Load.bStatus)
            ws_shiftable.write_number(row_shiftable, 6, Load.fPN)
            ws_shiftable.write_number(row_shiftable, 7, Load.fEtot)
            ws_shiftable.write_number(row_shiftable, 8, Load.Qhstart)
            ws_shiftable.write_number(row_shiftable, 9, Load.Qhend)
            ws_shiftable.write_number(row_shiftable, 10, Load.Qhstartmin)
            ws_shiftable.write_number(row_shiftable, 11, Load.Qhstartmax)        
            for h in range( 0, 96):
                ws_shiftable.write_number(row_shiftable, 12+h, Load.fPowerQh[0,h])  
        elif Load.sType == 'Elastic':            
            row_elastic = row_elastic + 1
            ws_elastic.write_string(row_elastic, 0, Load.sReference)        
            ws_elastic.write_string(row_elastic, 1, Load.sName)
            ws_elastic.write_string(row_elastic, 2, Load.sBuilding)
            ws_elastic.write_string(row_elastic, 3, Load.sSubzone)
            ws_elastic.write_string(row_elastic, 4, Load.sFacility)
            ws_elastic.write_string(row_elastic, 5, Load.sSubType)
            ws_elastic.write_number(row_elastic, 6, Load.fScale)
            ws_elastic.write_number(row_elastic, 7, Load.fScaleMin)        
            ws_elastic.write_number(row_elastic, 8, Load.bStatus)
            ws_elastic.write_number(row_elastic, 9, Load.fPN)
            ws_elastic.write_number(row_elastic, 10, Load.fEtot)
            for h in range( 0, 96):
                ws_elastic.write_number(row_elastic, 11+h, Load.fPowerQh[0,h])  
 
"""********************************************************"""


