# -*- coding: utf-8 -*-
"""
Created on Wed Oct 17 10:37:00 2018

@author: spauliuk
"""

import xlrd
import numpy as np
import matplotlib.pyplot as plt  
import pylab


#Sensitivity analysis folder order, by default, all strategies are off, one by one is implemented each at a time.
#Baseline (no RE)
#FabYieldImprovement
#EoL_RR_Improvement
#ChangeMaterialComposition
#ReduceMaterialContent
#ReUse_Materials
#LifeTimeExtension
#MoreIntenseUse


Region= 'G7'
Scope = 'G7 Vehicles'
FolderlistV_Sens =[
'G7_2019_4_25__18_21_20',
'G7_2019_4_26__7_32_5',
'G7_2019_4_26__7_39_6',
'G7_2019_4_26__7_43_59',
'G7_2019_4_26__7_48_5',
'G7_2019_4_26__7_51_24',
'G7_2019_4_26__7_55_26',
'G7_2019_4_26__7_58_48',
]

Scope = 'G7 Buildings'
FolderlistB_Sens =[
'G7_2019_4_25__18_49_22',
'G7_2019_4_26__8_2_20',
'G7_2019_4_26__8_6_35',
'G7_2019_4_26__8_9_49',
'G7_2019_4_26__8_13_53',
'G7_2019_4_26__8_18_5',
'G7_2019_4_26__8_18_5',
'G7_2019_4_26__8_21_33',
]


# Sensitivity plots

NS = 3
NR = 8

CumEmsV_Sens     = np.zeros((NS,NR)) # SSP-Scenario x RCP scenario x RES scenario
AnnEmsV2030_Sens = np.zeros((NS,NR)) # SSP-Scenario x RCP scenario x RES scenario
AnnEmsV2050_Sens = np.zeros((NS,NR)) # SSP-Scenario x RCP scenario x RES scenario

for r in range(0,NR): # RE scenario
    Path = 'C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + FolderlistV_Sens[r] + '\\'
    Resultfile = xlrd.open_workbook(Path + 'SysVar_TotalGHGFootprint.xls')
    Resultsheet = Resultfile.sheet_by_name('TotalGHGFootprint')
    for s in range(0,NS): # SSP scenario
        for t in range(0,35): # time
            CumEmsV_Sens[s,r] += Resultsheet.cell_value(t +2, 2*(s+1))
        AnnEmsV2030_Sens[s,r]  = Resultsheet.cell_value(16  , 2*(s+1))
        AnnEmsV2050_Sens[s,r]  = Resultsheet.cell_value(36  , 2*(s+1))
        
CumEmsB_Sens     = np.zeros((NS,NR)) # SSP-Scenario x RCP scenario x RES scenario
AnnEmsB2030_Sens = np.zeros((NS,NR)) # SSP-Scenario x RCP scenario x RES scenario
AnnEmsB2050_Sens = np.zeros((NS,NR)) # SSP-Scenario x RCP scenario x RES scenario

for r in range(0,NR): # RE scenario
    Path = 'C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + FolderlistB_Sens[r] + '\\'
    Resultfile = xlrd.open_workbook(Path + 'SysVar_TotalGHGFootprint.xls')
    Resultsheet = Resultfile.sheet_by_name('TotalGHGFootprint')
    for s in range(0,NS): # SSP scenario
        for t in range(0,35): # time
            CumEmsB_Sens[s,r] += Resultsheet.cell_value(t +2, 2*(s+1))
        AnnEmsB2030_Sens[s,r]  = Resultsheet.cell_value(16  , 2*(s+1))
        AnnEmsB2050_Sens[s,r]  = Resultsheet.cell_value(36  , 2*(s+1))
        
     

### Tornado plot for sensitivity
        
MyColorCycle = pylab.cm.Set1(np.arange(0,1,0.1)) # select 12 colors from the 'Paired' color map.            

Title = ['Passenger vehicles','residential buildings']
Scens = ['LED','SSP1','SSP2']
LWE   = ['higher yield, manuf. efficiency','EoL_RR_Improvement','Material substitution','ReduceMaterialContent','ReUse_Materials','LifeTimeExtension','MoreIntenseUse']

#2030 emissions

for m in range(0,NS): # SSP
    for n in range(0,2): # Veh/Buildings
        
        if n == 0:
            Data = AnnEmsV2030_Sens[:,1::]-np.einsum('S,n->Sn',AnnEmsV2030_Sens[:,0],np.ones(7))
            Base = AnnEmsV2030_Sens[:,0]
            
        if n == 1:
            Data = AnnEmsB2030_Sens[:,1::]-np.einsum('S,n->Sn',AnnEmsB2030_Sens[:,0],np.ones(7))
            Base = AnnEmsB2030_Sens[:,0]
            
        # plot results
    
        fig  = plt.figure(figsize=(5,8))
        ax1  = plt.axes([0.08,0.08,0.85,0.9])
            
        Poss = np.arange(7,0,-1)
        ax1.barh(Poss,Data[m,:], color=MyColorCycle, lw=0.4)

        # plot text and labels
        for mm in range(0,7):
            plt.text(Data[m,:].min()*0.9, 7.3-mm, LWE[mm],fontsize=14,fontweight='bold')          
            plt.text(15, 7.3-mm, ("%3.0f" % Data[m,mm]),fontsize=14,fontweight='bold')          
        
        plt.text(Data[m,:].min()*0.55, 7.8, 'Baseline: ' + ("%3.0f" % Base[m]) + ' Mt/yr.',fontsize=14,fontweight='bold')

        plt.title('2030 GHG emissions, sensitivity, ' + Region + '_ ' + Title[n] + '_' + Scens[m] + '.', fontsize = 18)
        plt.ylabel('RE strategies.', fontsize = 18)
        plt.xlabel('Mt of CO2-eq.', fontsize = 14)
        #plt.xticks([0.25,1.25,2.25,3.25,4.25,5.25])
        plt.yticks([])
        #ax1.set_xticklabels([], rotation =90, fontsize = 21, fontweight = 'normal')
        #plt_lgd  = plt.legend(handles = ProxyHandlesList,labels = LWE,shadow = False, prop={'size':16},ncol=1, loc = 'upper right' ,bbox_to_anchor=(1.91, 1)) 
        plt.axis([Data[m,:].min() -15, Data[m,:].max() +80, 0.7, 8.1])
    
        plt.show()
        fig_name = '2030_GHG_Sens_' + Region + '_ ' + Title[n] + '_' + Scens[m] + '.png'
        fig.savefig('C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + fig_name, dpi = 400, bbox_inches='tight')             
            
        fig  = plt.figure(figsize=(5,8))
        ax1  = plt.axes([0.08,0.08,0.85,0.9])
            
        Poss = np.arange(7,0,-1)
        ax1.barh(Poss,Data[m,:], color=MyColorCycle, lw=0.4)

        # plot text and labels
        for mm in range(0,7):
            plt.text(Data[m,:].min()*0.9, 7.3-mm, LWE[mm] + ': ' + ("%3.0f" % Data[m,mm]),fontsize=14,fontweight='bold')          
        
        plt.text(Data[m,:].min()*0.55, 7.8, 'Baseline: ' + ("%3.0f" % Base[m]) + ' Mt/yr.',fontsize=14,fontweight='bold')

        plt.title('2030 GHG emissions, sensitivity, ' + Region + '_ ' + Title[n] + '_' + Scens[m] + '.', fontsize = 18)
        plt.ylabel('RE strategies.', fontsize = 18)
        plt.xlabel('Mt of CO2-eq.', fontsize = 14)
        #plt.xticks([0.25,1.25,2.25,3.25,4.25,5.25])
        plt.yticks([])
        #ax1.set_xticklabels([], rotation =90, fontsize = 21, fontweight = 'normal')
        #plt_lgd  = plt.legend(handles = ProxyHandlesList,labels = LWE,shadow = False, prop={'size':16},ncol=1, loc = 'upper right' ,bbox_to_anchor=(1.91, 1)) 
        plt.axis([Data[m,:].min() -15, Data[m,:].max() +80, 0.7, 8.1])
    
        plt.show()
        fig_name = '2030_GHG_Sens_' + Region + '_ ' + Title[n] + '_' + Scens[m] + '_V2.png'
        fig.savefig('C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + fig_name, dpi = 400, bbox_inches='tight')             
      
        
#2050 emissions

for m in range(0,NS): # SSP
    for n in range(0,2): # Veh/Buildings
        
        if n == 0:
            Data = AnnEmsV2050_Sens[:,1::]-np.einsum('S,n->Sn',AnnEmsV2050_Sens[:,0],np.ones(7))
            Base = AnnEmsV2050_Sens[:,0]
            
        if n == 1:
            Data = AnnEmsB2050_Sens[:,1::]-np.einsum('S,n->Sn',AnnEmsB2050_Sens[:,0],np.ones(7))
            Base = AnnEmsB2050_Sens[:,0]
            
        # plot results
    
        fig  = plt.figure(figsize=(5,8))
        ax1  = plt.axes([0.08,0.08,0.85,0.9])
            
        Poss = np.arange(7,0,-1)
        ax1.barh(Poss,Data[m,:], color=MyColorCycle, lw=0.4)

        # plot text and labels
        for mm in range(0,7):
            plt.text(Data[m,:].min()*0.9, 7.3-mm, LWE[mm],fontsize=14,fontweight='bold')          
            plt.text(15, 7.3-mm, ("%3.0f" % Data[m,mm]),fontsize=14,fontweight='bold')          
        
        plt.text(Data[m,:].min()*0.55, 7.8, 'Baseline: ' + ("%3.0f" % Base[m]) + ' Mt/yr.',fontsize=14,fontweight='bold')

        plt.title('2050 GHG emissions, sensitivity, ' + Region + '_ ' + Title[n] + '_' + Scens[m] + '.', fontsize = 18)
        plt.ylabel('RE strategies.', fontsize = 18)
        plt.xlabel('Mt of CO2-eq.', fontsize = 14)
        #plt.xticks([0.25,1.25,2.25,3.25,4.25,5.25])
        plt.yticks([])
        #ax1.set_xticklabels([], rotation =90, fontsize = 21, fontweight = 'normal')
        #plt_lgd  = plt.legend(handles = ProxyHandlesList,labels = LWE,shadow = False, prop={'size':16},ncol=1, loc = 'upper right' ,bbox_to_anchor=(1.91, 1)) 
        plt.axis([Data[m,:].min() -15, Data[m,:].max() +80, 0.7, 8.1])
    
        plt.show()
        fig_name = '2050_GHG_Sens_' + Region + '_ ' + Title[n] + '_' + Scens[m] + '.png'
        fig.savefig('C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + fig_name, dpi = 400, bbox_inches='tight')             
            
        fig  = plt.figure(figsize=(5,8))
        ax1  = plt.axes([0.08,0.08,0.85,0.9])
            
        Poss = np.arange(7,0,-1)
        ax1.barh(Poss,Data[m,:], color=MyColorCycle, lw=0.4)

        # plot text and labels
        for mm in range(0,7):
            plt.text(Data[m,:].min()*0.9, 7.3-mm, LWE[mm] + ': ' + ("%3.0f" % Data[m,mm]),fontsize=14,fontweight='bold')          
        
        plt.text(Data[m,:].min()*0.55, 7.8, 'Baseline: ' + ("%3.0f" % Base[m]) + ' Mt/yr.',fontsize=14,fontweight='bold')

        plt.title('2050 GHG emissions, sensitivity, ' + Region + '_ ' + Title[n] + '_' + Scens[m] + '.', fontsize = 18)
        plt.ylabel('RE strategies.', fontsize = 18)
        plt.xlabel('Mt of CO2-eq.', fontsize = 14)
        #plt.xticks([0.25,1.25,2.25,3.25,4.25,5.25])
        plt.yticks([])
        #ax1.set_xticklabels([], rotation =90, fontsize = 21, fontweight = 'normal')
        #plt_lgd  = plt.legend(handles = ProxyHandlesList,labels = LWE,shadow = False, prop={'size':16},ncol=1, loc = 'upper right' ,bbox_to_anchor=(1.91, 1)) 
        plt.axis([Data[m,:].min() -15, Data[m,:].max() +80, 0.7, 8.1])
    
        plt.show()
        fig_name = '2050_GHG_Sens_' + Region + '_ ' + Title[n] + '_' + Scens[m] + '_V2.png'
        fig.savefig('C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + fig_name, dpi = 400, bbox_inches='tight')             
      
#2050 cum. emissions

for m in range(0,NS): # SSP
    for n in range(0,2): # Veh/Buildings
        
        if n == 0:
            Data = CumEmsV_Sens[:,1::]-np.einsum('S,n->Sn',CumEmsV_Sens[:,0],np.ones(7))
            Base = CumEmsV_Sens[:,0]
        if n == 1:
            Data = CumEmsB_Sens[:,1::]-np.einsum('S,n->Sn',CumEmsB_Sens[:,0],np.ones(7))
            Base = CumEmsB_Sens[:,0]
    
        # plot results
    
        fig  = plt.figure(figsize=(5,8))
        ax1  = plt.axes([0.08,0.08,0.85,0.9])
            
        Poss = np.arange(7,0,-1)

        ax1.barh(Poss,Data[m,:], color=MyColorCycle, lw=0.4)
        
        # plot text and labels
        for mm in range(0,7):
            plt.text(Data[m,:].min() *0.9, 7.3-mm, LWE[mm],fontsize=14,fontweight='bold')          
            plt.text(-Data[m,:].min() *0.1, 7.3-mm, ("%2.0f" % Data[m,mm]),fontsize=14,fontweight='bold')          

        plt.text(Data[m,:].min()*0.7, 7.8, 'Baseline: ' + ("%2.0f" % Base[m]) + ' Mt/yr.',fontsize=14,fontweight='bold')
        plt.title('2016-2050 cum. GHG, sensitivity, ' + Region + '_ ' + Title[n] + '_' + Scens[m] + '.', fontsize = 18)
        plt.ylabel('RE strategies.', fontsize = 18)
        plt.xlabel('Mt of CO2-eq.', fontsize = 14)
        #plt.xticks([0.25,1.25,2.25,3.25,4.25,5.25])
        plt.yticks([])
        #ax1.set_xticklabels([], rotation =90, fontsize = 21, fontweight = 'normal')
        #plt_lgd  = plt.legend(handles = ProxyHandlesList,labels = LWE,shadow = False, prop={'size':16},ncol=1, loc = 'upper right' ,bbox_to_anchor=(1.91, 1)) 
        plt.axis([Data[m,:].min() *1.05, -Data[m,:].min() *0.05, 0.7, 8.1])
    
        plt.show()
        fig_name = 'Cum_GHG_Sens_' + Region + '_ ' + Title[n] + '_' + Scens[m] + '.png'
        fig.savefig('C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + fig_name, dpi = 400, bbox_inches='tight')             
            


