# -*- coding: utf-8 -*-
"""
Created on Wed Oct 17 10:37:00 2018

@author: spauliuk
"""

import xlrd
import numpy as np
import matplotlib.pyplot as plt  
import pylab

# FileOrder:
# 1) None
# 2) + EoL + FYI
# 3) + EoL + FYI + LWE + MSu
# 4) + EoL + FYI + LWE + MSu + ReU (+LTE)
# 5) + EoL + FYI + LWE + MSu + ReU (+LTE) + MIU = ALL


Region= 'Canada'
Scope = 'Canada_Vehicles'
FolderlistV =[
'Canada_2019_4_18__19_37_51',
'Canada_2019_4_18__19_38_53',
'Canada_2019_4_18__19_40_4',
'Canada_2019_4_18__19_41_8',
'Canada_2019_4_18__19_42_14'
]

Scope = 'Canada_Buildings'
FolderlistB =[
'Canada_2019_4_18__19_44_29',
'Canada_2019_4_18__19_45_31',
'Canada_2019_4_18__19_46_33',
'Canada_2019_4_18__19_47_36',
'Canada_2019_4_18__19_48_40'
]


Region= 'USA'
Scope = 'USA_Vehicles'
FolderlistV =[
'USA_2019_4_18__19_51_23',
'USA_2019_4_18__19_52_53',
'USA_2019_4_18__19_54_18',
'USA_2019_4_18__19_55_20',
'USA_2019_4_18__19_56_25'
]

Scope = 'USA_buildings'
FolderlistB =[
'USA_2019_4_18__20_2_49',
'USA_2019_4_18__20_1_43',
'USA_2019_4_18__20_0_22',
'USA_2019_4_18__19_58_59',
'USA_2019_4_18__19_57_34'
]


Region= 'France'
Scope = 'France_Vehicles'
FolderlistV =[
'France_2019_4_18__20_4_52',
'France_2019_4_18__20_6_10',
'France_2019_4_18__20_7_12',
'France_2019_4_18__20_8_28',
'France_2019_4_18__20_9_43'
]

Scope = 'France_Buildings'
FolderlistB =[
'France_2019_4_18__20_10_46',
'France_2019_4_18__20_12_23',
'France_2019_4_18__20_14_8',
'France_2019_4_18__20_18_29',
'France_2019_4_18__20_19_46'
]


Region= 'Germany'
Scope = 'Germany_Vehicles'
FolderlistV =[
'Germany_2019_4_18__19_17_19',
'Germany_2019_4_18__19_16_17',
'Germany_2019_4_18__19_15_14',
'Germany_2019_4_18__19_13_48',
'Germany_2019_4_18__19_12_41'
]

Scope = 'Germany_Buildings'
FolderlistB =[
'Germany_2019_4_18__19_4_5',
'Germany_2019_4_18__19_5_32',
'Germany_2019_4_18__19_6_35',
'Germany_2019_4_18__19_7_37',
'Germany_2019_4_18__19_8_46'
]


Region= 'Japan'
Scope = 'Japan_Vehicles'
FolderlistV =[
'Japan_2019_4_18__21_12_0',
'Japan_2019_4_18__21_15_30',
'Japan_2019_4_18__21_16_56',
'Japan_2019_4_18__21_18_10',
'Japan_2019_4_18__21_19_18'
]

Scope = 'Japan_Buildings'
FolderlistB =[
'Japan_2019_4_18__21_20_40',
'Japan_2019_4_18__21_21_45',
'Japan_2019_4_18__21_23_0',
'Japan_2019_4_18__21_24_4',
'Japan_2019_4_18__21_25_7'
]


Region= 'Italy'
Scope = 'Italy_Vehicles'
FolderlistV =[
'Italy_2019_4_18__20_24_58',
'Italy_2019_4_18__20_57_50',
'Italy_2019_4_18__20_59_46',
'Italy_2019_4_18__21_1_6',
'Italy_2019_4_18__21_2_34'
]

Scope = 'Italy_Buildings'
FolderlistB =[
'Italy_2019_4_18__21_3_40',
'Italy_2019_4_18__21_4_54',
'Italy_2019_4_18__21_6_29',
'Italy_2019_4_18__21_7_46',
'Italy_2019_4_18__21_8_57'
]


Region= 'UK'
Scope = 'UK_Vehicles'
FolderlistV =[
'UK_2019_4_18__21_27_39',
'UK_2019_4_18__21_31_44',
'UK_2019_4_18__21_33_8',
'UK_2019_4_18__21_35_11',
'UK_2019_4_18__21_36_29'
]

Scope = 'UK_Buildings'
FolderlistB =[
'UK_2019_4_18__21_37_33',
'UK_2019_4_18__21_38_39',
'UK_2019_4_18__21_40_3',
'UK_2019_4_18__21_41_36',
'UK_2019_4_18__21_43_8'
]


Region= 'G7'
Scope = 'G7 Vehicles'
FolderlistV =[
'G7_2019_4_18__22_8_8',
'G7_2019_4_18__22_13_30',
'G7_2019_4_18__22_16_48',
'G7_2019_4_18__22_19_48',
'G7_2019_4_18__22_23_47'
]

Scope = 'G7 Buildings'
FolderlistB =[
'G7_2019_4_18__22_27_0',
'G7_2019_4_18__22_31_32',
'G7_2019_4_18__22_35_15',
'G7_2019_4_18__22_39_25',
'G7_2019_4_18__22_43_45'
]








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
'G7_2019_4_18__22_8_8',
'G7_2019_4_19__7_55_16',
'G7_2019_4_19__8_0_41',
'G7_2019_4_19__8_9_6',
'G7_2019_4_19__8_13_58',
'G7_2019_4_19__8_19_17',
'G7_2019_4_19__8_22_28',
'G7_2019_4_19__8_27_9',
]

Scope = 'G7 Buildings'
FolderlistB_Sens =[
'G7_2019_4_18__22_27_0',
'G7_2019_4_19__8_47_52',
'G7_2019_4_19__9_24_11',
'G7_2019_4_19__9_27_35',
'G7_2019_4_19__9_32_20',
'G7_2019_4_19__9_35_45',
'G7_2019_4_19__9_42_3',
'G7_2019_4_19__9_45_20',
]


# Waterfall plots.

NS = 3
NC = 2
NR = 5

CumEmsV     = np.zeros((NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario
AnnEmsV2030 = np.zeros((NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario
AnnEmsV2050 = np.zeros((NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario
ASummaryV    = np.zeros((9,NR)) # For direct copy-paste to Excel

for r in range(0,NR): # RE scenario
    Path = 'C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + FolderlistV[r] + '\\'
    Resultfile = xlrd.open_workbook(Path + 'SysVar_TotalGHGFootprint.xls')
    Resultsheet = Resultfile.sheet_by_name('TotalGHGFootprint')
    for s in range(0,NS): # SSP scenario
        for c in range(0,NC):
            for t in range(0,35): # time
                CumEmsV[s,c,r] += Resultsheet.cell_value(t +2, 1 + c + NC*s)
            AnnEmsV2030[s,c,r]  = Resultsheet.cell_value(16  , 1 + c + NC*s)
            AnnEmsV2050[s,c,r]  = Resultsheet.cell_value(36  , 1 + c + NC*s)
            
ASummaryV[0:3,:] = AnnEmsV2030[:,1,:].copy()
ASummaryV[3:6,:] = AnnEmsV2050[:,1,:].copy()
ASummaryV[6::,:] = CumEmsV[:,1,:].copy()
        

CumEmsB     = np.zeros((NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario
AnnEmsB2030 = np.zeros((NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario
AnnEmsB2050 = np.zeros((NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario
ASummaryB    = np.zeros((9,NR)) # For direct copy-paste to Excel

for r in range(0,NR): # RE scenario
    Path = 'C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + FolderlistB[r] + '\\'
    Resultfile = xlrd.open_workbook(Path + 'SysVar_TotalGHGFootprint.xls')
    Resultsheet = Resultfile.sheet_by_name('TotalGHGFootprint')
    for s in range(0,NS): # SSP scenario
        for c in range(0,NC):
            for t in range(0,35): # time
                CumEmsB[s,c,r] += Resultsheet.cell_value(t +2, 1 + c + NC*s)
            AnnEmsB2030[s,c,r]  = Resultsheet.cell_value(16  , 1 + c + NC*s)
            AnnEmsB2050[s,c,r]  = Resultsheet.cell_value(36  , 1 + c + NC*s)
            
ASummaryB[0:3,:] = AnnEmsB2030[:,1,:].copy()
ASummaryB[3:6,:] = AnnEmsB2050[:,1,:].copy()
ASummaryB[6::,:] = CumEmsB[:,1,:].copy()
        
# Waterfall plot            
MyColorCycle = pylab.cm.Set1(np.arange(0,1,0.1)) # select 12 colors from the 'Paired' color map.            

Title = ['Passenger vehicles','residential buildings']
Scens = ['LED','SSP1','SSP2']
LWE   = ['No RE','recycling improvemt.','light-weighting','re-use','More intense use','All RE stratgs.']

# Cumulative emissions
for m in range(0,NS): # SSP
    for n in range(0,2): # Veh/Buildings
        
        if n == 0:
            Data = np.einsum('SR->RS',CumEmsV[:,1,:])
        if n == 1:
            Data = np.einsum('SR->RS',CumEmsB[:,1,:])
    
        inc = -100 * (Data[0,m] - Data[4,m])/Data[0,m]
    
        Left  = Data[0,m]
        Right = Data[4,m]
        # plot results
        bw = 0.5
        ga = 0.3
    
        fig  = plt.figure(figsize=(5,8))
        ax1  = plt.axes([0.08,0.08,0.85,0.9])
    
        ProxyHandlesList = []   # For legend     
        # plot bars for domestic footprint
        ax1.fill_between([0,0+bw], [0,0],[Left,Left],linestyle = '--', facecolor =MyColorCycle[0,:], linewidth = 0.0)
        ax1.fill_between([1,1+bw], [Data[1,m],Data[1,m]],[Left,Left],linestyle = '--', facecolor =MyColorCycle[1,:], linewidth = 0.0)
        ax1.fill_between([2,2+bw], [Data[2,m],Data[2,m]],[Data[1,m],Data[1,m]],linestyle = '--', facecolor =MyColorCycle[2,:], linewidth = 0.0)
        ax1.fill_between([3,3+bw], [Data[3,m],Data[3,m]],[Data[2,m],Data[2,m]],linestyle = '--', facecolor =MyColorCycle[3,:], linewidth = 0.0)
        ax1.fill_between([4,4+bw], [Data[4,m],Data[4,m]],[Data[3,m],Data[3,m]],linestyle = '--', facecolor =MyColorCycle[4,:], linewidth = 0.0)
        ax1.fill_between([5,5+bw], [0,0],[Data[4,m],Data[4,m]],linestyle = '--', facecolor =MyColorCycle[5,:], linewidth = 0.0)
        
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[0,:])) # create proxy artist for legend
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[1,:])) # create proxy artist for legend
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[2,:])) # create proxy artist for legend
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[3,:])) # create proxy artist for legend
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[4,:])) # create proxy artist for legend
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[5,:])) # create proxy artist for legend
        
        # plot lines:
        plt.plot([0,5.5],[Left,Left],linestyle = '-', linewidth = 0.5, color = 'k')
        plt.plot([1,2.5],[Data[1,m],Data[1,m]],linestyle = '-', linewidth = 0.5, color = 'k')
        plt.plot([2,3.5],[Data[2,m],Data[2,m]],linestyle = '-', linewidth = 0.5, color = 'k')
        plt.plot([3,4.5],[Data[3,m],Data[3,m]],linestyle = '-', linewidth = 0.5, color = 'k')
        plt.plot([4,5.5],[Data[4,m],Data[4,m]],linestyle = '-', linewidth = 0.5, color = 'k')

        plt.arrow(5.25, Data[4,m],0, Data[0,m]-Data[4,m], lw = 0.8, ls = '-', shape = 'full',
              length_includes_head = True, head_width =0.1, head_length =0.01*Left, ec = 'k', fc = 'k')
        plt.arrow(5.25,Data[0,m],0,Data[4,m]-Data[0,m], lw = 0.8, ls = '-', shape = 'full',
              length_includes_head = True, head_width =0.1, head_length =0.01*Left, ec = 'k', fc = 'k')

        # plot text and labels
        plt.text(3.85, 0.94 *Left, ("%3.0f" % inc) + ' %',fontsize=18,fontweight='bold')          
        plt.text(2.3, 1.007 *Left, Scens[m],fontsize=18,fontweight='bold') 
        plt.title('RE strategies and GHG emissions, ' + Title[n] + '.', fontsize = 18)
        plt.ylabel('Cumulative GHG emissions 2016-2050, Mt.', fontsize = 18)
        plt.xticks([0.25,1.25,2.25,3.25,4.25,5.25])
        plt.yticks(fontsize =18)
        ax1.set_xticklabels([], rotation =90, fontsize = 21, fontweight = 'normal')
        plt_lgd  = plt.legend(handles = ProxyHandlesList,labels = LWE,shadow = False, prop={'size':16},ncol=1, loc = 'upper right' ,bbox_to_anchor=(1.91, 1)) 
        plt.axis([-0.2, 5.7, 0.9*Right, 1.02*Left])
    
        plt.show()
        fig_name = 'Cum_GHG_' + Region + '_ ' + Title[n] + '_' + Scens[m] + '.png'
        fig.savefig('C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + fig_name, dpi = 400, bbox_inches='tight')             
            



# 2050 emissions
for m in range(0,NS): # SSP
    for n in range(0,2): # Veh/Buldings
        
        if n == 0:
            Data = np.einsum('SR->RS',AnnEmsV2050[:,1,:])
        if n == 1:
            Data = np.einsum('SR->RS',AnnEmsB2050[:,1,:])
    
        inc = -100 * (Data[0,m] - Data[4,m])/Data[0,m]
    
        Left  = Data[0,m]
        Right = Data[4,m]
        # plot results
        bw = 0.5
        ga = 0.3
    
        fig  = plt.figure(figsize=(5,8))
        ax1  = plt.axes([0.08,0.08,0.85,0.9])
    
        ProxyHandlesList = []   # For legend     
        # plot bars for domestic footprint
        ax1.fill_between([0,0+bw], [0,0],[Left,Left],linestyle = '--', facecolor =MyColorCycle[0,:], linewidth = 0.0)
        ax1.fill_between([1,1+bw], [Data[1,m],Data[1,m]],[Left,Left],linestyle = '--', facecolor =MyColorCycle[1,:], linewidth = 0.0)
        ax1.fill_between([2,2+bw], [Data[2,m],Data[2,m]],[Data[1,m],Data[1,m]],linestyle = '--', facecolor =MyColorCycle[2,:], linewidth = 0.0)
        ax1.fill_between([3,3+bw], [Data[3,m],Data[3,m]],[Data[2,m],Data[2,m]],linestyle = '--', facecolor =MyColorCycle[3,:], linewidth = 0.0)
        ax1.fill_between([4,4+bw], [Data[4,m],Data[4,m]],[Data[3,m],Data[3,m]],linestyle = '--', facecolor =MyColorCycle[4,:], linewidth = 0.0)
        ax1.fill_between([5,5+bw], [0,0],[Data[4,m],Data[4,m]],linestyle = '--', facecolor =MyColorCycle[5,:], linewidth = 0.0)
        
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[0,:])) # create proxy artist for legend
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[1,:])) # create proxy artist for legend
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[2,:])) # create proxy artist for legend
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[3,:])) # create proxy artist for legend
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[4,:])) # create proxy artist for legend
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[5,:])) # create proxy artist for legend
        
        # plot lines:
        plt.plot([0,5.5],[Left,Left],linestyle = '-', linewidth = 0.5, color = 'k')
        plt.plot([1,2.5],[Data[1,m],Data[1,m]],linestyle = '-', linewidth = 0.5, color = 'k')
        plt.plot([2,3.5],[Data[2,m],Data[2,m]],linestyle = '-', linewidth = 0.5, color = 'k')
        plt.plot([3,4.5],[Data[3,m],Data[3,m]],linestyle = '-', linewidth = 0.5, color = 'k')
        plt.plot([4,5.5],[Data[4,m],Data[4,m]],linestyle = '-', linewidth = 0.5, color = 'k')

        plt.arrow(5.25, Data[4,m],0, Data[0,m]-Data[4,m], lw = 0.8, ls = '-', shape = 'full',
              length_includes_head = True, head_width =0.1, head_length =0.01*Left, ec = 'k', fc = 'k')
        plt.arrow(5.25,Data[0,m],0,Data[4,m]-Data[0,m], lw = 0.8, ls = '-', shape = 'full',
              length_includes_head = True, head_width =0.1, head_length =0.01*Left, ec = 'k', fc = 'k')

        # plot text and labels
        plt.text(3.85, 0.94 *Left, ("%3.0f" % inc) + ' %',fontsize=18,fontweight='bold')          
        plt.text(2.3, 1.007 *Left, Scens[m],fontsize=18,fontweight='bold') 
        plt.title('RE strategies and GHG emissions, ' + Title[n] + '.', fontsize = 18)
        plt.ylabel('2050 GHG emissions, Mt.', fontsize = 18)
        plt.xticks([0.25,1.25,2.25,3.25,4.25,5.25])
        plt.yticks(fontsize =18)
        ax1.set_xticklabels([], rotation =90, fontsize = 21, fontweight = 'normal')
        plt_lgd  = plt.legend(handles = ProxyHandlesList,labels = LWE,shadow = False, prop={'size':16},ncol=1, loc = 'upper right' ,bbox_to_anchor=(1.91, 1)) 
        plt.axis([-0.2, 5.7, 0.7*Right, 1.025*Left])
    
        plt.show()
        fig_name = 'GHG_2050_' + Region + '_ ' + Title[n] + '_' + Scens[m] + '.png'
        fig.savefig('C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + fig_name, dpi = 400, bbox_inches='tight')             



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
LWE   = ['FabYieldImprovement','EoL_RR_Improvement','ChangeMaterialCompos.','ReduceMaterialContent','ReUse_Materials','LifeTimeExtension','MoreIntenseUse']

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
            





### Area plot RE
        
        
# Select scenario list: same as for bar chart above
# E.g. for the USA, run code lines 41 to 59.

NS = 3
NC = 2
NR = 5
Nt = 35

AnnEmsV = np.zeros((Nt,NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario
AnnEmsB = np.zeros((Nt,NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario

for r in range(0,NR): # RE scenario
    Path = 'C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + FolderlistV[r] + '\\'
    Resultfile = xlrd.open_workbook(Path + 'SysVar_TotalGHGFootprint.xls')
    Resultsheet = Resultfile.sheet_by_name('TotalGHGFootprint')
    for s in range(0,NS): # SSP scenario
        for c in range(0,NC):
            for t in range(0,35): # time
                AnnEmsV[t,s,c,r] = Resultsheet.cell_value(t +2, 1 + c + NC*s)

for r in range(0,NR): # RE scenario
    Path = 'C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + FolderlistB[r] + '\\'
    Resultfile = xlrd.open_workbook(Path + 'SysVar_TotalGHGFootprint.xls')
    Resultsheet = Resultfile.sheet_by_name('TotalGHGFootprint')
    for s in range(0,NS): # SSP scenario
        for c in range(0,NC):
            for t in range(0,35): # time
                AnnEmsB[t,s,c,r] = Resultsheet.cell_value(t +2, 1 + c + NC*s)
        
# Area plot, stacked
MyColorCycle = pylab.cm.Set1(np.arange(0,1,0.1)) # select 12 colors from the 'Paired' color map.            
grey0_9      = np.array([0.9,0.9,0.9,1])

Title      = ['Passenger vehicles','residential buildings']
Scens      = ['LED','SSP1','SSP2']
LWE_area   = ['recycling improvement','light-weighting','re-use','more intense use']     

#mS = 1
#mR = 1
mRCP = 1 # select RCP2.6, which has full implementation of RE strategies by 2050.
for mS in range(0,NS): # SSP
    for mR in range(0,2): # Veh/Buildings
        
        if mR == 0:
            Data = AnnEmsV[:,mS,mRCP,:]
        if mR == 1:
            Data = AnnEmsB[:,mS,mRCP,:]
    
    
        fig  = plt.figure(figsize=(8,5))
        ax1  = plt.axes([0.08,0.08,0.85,0.9])
        
        ProxyHandlesList = []   # For legend     
        
        # plot bars for domestic footprint
        ax1.fill_between(np.arange(2016,2051),np.zeros((Nt)), Data[:,-1], linestyle = '-', facecolor = grey0_9, linewidth = 1.0)
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=grey0_9)) # create proxy artist for legend
        for m in range(4,0,-1):
            ax1.fill_between(np.arange(2016,2051),Data[:,m], Data[:,m-1], linestyle = '-', facecolor = MyColorCycle[2*m,:], linewidth = 1.0)
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[2*m,:])) # create proxy artist for legend
            
        #plt.text(Data[m,:].min()*0.55, 7.8, 'Baseline: ' + ("%3.0f" % Base[m]) + ' Mt/yr.',fontsize=14,fontweight='bold')
        
        plt.title('GHG emissions, stacked by RE strategy, \n' + Region + ', ' + Title[mR] + ', ' + Scens[mS] + '.', fontsize = 18)
        plt.ylabel('Mt of CO2-eq.', fontsize = 18)
        plt.xlabel('Year', fontsize = 18)
        plt.xticks(fontsize=18)
        plt.yticks(fontsize=18)
        if mR == 0: # vehicles, legend lower left
            plt_lgd  = plt.legend(handles = reversed(ProxyHandlesList),labels = LWE_area, shadow = False, prop={'size':16},ncol=1, loc = 'lower left')# ,bbox_to_anchor=(1.91, 1)) 
        if mR == 1: # buildings, legend upper right
            plt_lgd  = plt.legend(handles = reversed(ProxyHandlesList),labels = LWE_area, shadow = False, prop={'size':16},ncol=1, loc = 'upper right')# ,bbox_to_anchor=(1.91, 1)) 
        ax1.set_xlim([2015, 2051])
        
        plt.show()
        fig_name = 'GHG_TimeSeries_Stacked_' + Region + '_ ' + Title[mR] + '_' + Scens[mS] + '.png'
        fig.savefig('C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + fig_name, dpi = 400, bbox_inches='tight')             






           
##### line Plot overview of primary steel and steel recycling

# Select scenario list: same as for bar chart above
# E.g. for the USA, run code lines 41 to 59.

MyColorCycle = pylab.cm.Paired(np.arange(0,1,0.2))
#linewidth = [1.2,2.4,1.2,1.2,1.2]
linewidth  = [1.2,2,1.2]
linewidth2 = [1.2,2,1.2]

Figurecounter = 1
ColorOrder         = [1,0,3]
        
NS = 3
NC = 2
NR = 5
Nt = 35

# Primary steel
AnnEmsV_PrimarySteel = np.zeros((Nt,NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario
AnnEmsB_PrimarySteel = np.zeros((Nt,NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario

for r in range(0,NR): # RE scenario
    Path = 'C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + FolderlistV[r] + '\\'
    Resultfile1  = xlrd.open_workbook(Path + 'SysVar_TotalGHGFootprint.xls')
    Resultsheet1 = Resultfile1.sheet_by_name('Cover')
    UUID         = Resultsheet1.cell_value(3,2)
    Resultfile2  = xlrd.open_workbook(Path + 'ODYM_RECC_ModelResults_' + UUID + '.xls')
    Resultsheet2 = Resultfile2.sheet_by_name('Model_Results')
    for s in range(0,NS): # SSP scenario
        for c in range(0,NC):
            for t in range(0,35): # time
                AnnEmsV_PrimarySteel[t,s,c,r] = Resultsheet2.cell_value(19+ 2*s +c,t+8)
                
for r in range(0,NR): # RE scenario
    Path = 'C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + FolderlistB[r] + '\\'
    Resultfile1  = xlrd.open_workbook(Path + 'SysVar_TotalGHGFootprint.xls')
    Resultsheet1 = Resultfile1.sheet_by_name('Cover')
    UUID         = Resultsheet1.cell_value(3,2)
    Resultfile2  = xlrd.open_workbook(Path + 'ODYM_RECC_ModelResults_' + UUID + '.xls')
    Resultsheet2 = Resultfile2.sheet_by_name('Model_Results')
    for s in range(0,NS): # SSP scenario
        for c in range(0,NC):
            for t in range(0,35): # time
                AnnEmsB_PrimarySteel[t,s,c,r] = Resultsheet2.cell_value(19+ 2*s +c,t+8)                
                
Title      = ['Passenger vehicles','residential buildings']
ScensL     = ['SSP2, no REFs','SSP2, full REF spectrum','SSP1, no REFs','SSP1, full REF spectrum','LED, no REFs','LED, full REF spectrum']

#mS = 1
#mR = 1
mRCP = 1 # select RCP2.6, which has full implementation of RE strategies by 2050.
for mR in range(0,2): # Veh/Buildings
    
    if mR == 0:
        Data = AnnEmsV_PrimarySteel[:,:,mRCP,:]
    if mR == 1:
        Data = AnnEmsB_PrimarySteel[:,:,mRCP,:]


    fig  = plt.figure(figsize=(8,5))
    ax1  = plt.axes([0.08,0.08,0.85,0.9])
    
    ProxyHandlesList = []   # For legend     
    
    for mS in range(NS-1,-1,-1):
        ax1.plot(np.arange(2016,2051), Data[:,mS,0],  linewidth = linewidth[mS],  linestyle = '-',  color = MyColorCycle[ColorOrder[mS],:])
        #ProxyHandlesList.append(plt.line((0, 0), 1, 1, fc=MyColorCycle[m,:]))
        ax1.plot(np.arange(2016,2051), Data[:,mS,-1], linewidth = linewidth2[mS], linestyle = '--', color = MyColorCycle[ColorOrder[mS],:])
        #ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))     
    plt_lgd  = plt.legend(ScensL,shadow = False, prop={'size':14}, loc = 'upper left',bbox_to_anchor=(1.05, 1))    
    plt.ylabel('Primary steel and iron, Mt/yr.', fontsize = 18) 
    plt.xlabel('year', fontsize = 18)         
    plt.title('Primary steel, by socio-economic scenario, \n' + Region + ', ' + Title[mR] + '.', fontsize = 18)
    plt.xticks(fontsize=18)
    plt.yticks(fontsize=18)
    ax1.set_xlim([2015, 2051])
    plt.gca().set_ylim(bottom=0)
    
    plt.show()
    fig_name = 'PrimarySteel_TimeSeries_' + Region + '_ ' + Title[mR] + '.png'
    fig.savefig('C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + fig_name, dpi = 400, bbox_inches='tight')             


# recycled steel (both used within sector and exported to other sectors)
AnnEmsV_SecondarySteel = np.zeros((Nt,NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario
AnnEmsB_SecondarySteel = np.zeros((Nt,NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario

for r in range(0,NR): # RE scenario
    Path = 'C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + FolderlistV[r] + '\\'
    Resultfile1  = xlrd.open_workbook(Path + 'SysVar_TotalGHGFootprint.xls')
    Resultsheet1 = Resultfile1.sheet_by_name('Cover')
    UUID         = Resultsheet1.cell_value(3,2)
    Resultfile2  = xlrd.open_workbook(Path + 'ODYM_RECC_ModelResults_' + UUID + '.xls')
    Resultsheet2 = Resultfile2.sheet_by_name('Model_Results')
    for s in range(0,NS): # SSP scenario
        for c in range(0,NC):
            for t in range(0,35): # time
                AnnEmsV_SecondarySteel[t,s,c,r] = Resultsheet2.cell_value(151+ 2*s +c,t+8)
                
for r in range(0,NR): # RE scenario
    Path = 'C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + FolderlistB[r] + '\\'
    Resultfile1  = xlrd.open_workbook(Path + 'SysVar_TotalGHGFootprint.xls')
    Resultsheet1 = Resultfile1.sheet_by_name('Cover')
    UUID         = Resultsheet1.cell_value(3,2)
    Resultfile2  = xlrd.open_workbook(Path + 'ODYM_RECC_ModelResults_' + UUID + '.xls')
    Resultsheet2 = Resultfile2.sheet_by_name('Model_Results')
    for s in range(0,NS): # SSP scenario
        for c in range(0,NC):
            for t in range(0,35): # time
                AnnEmsB_SecondarySteel[t,s,c,r] = Resultsheet2.cell_value(151+ 2*s +c,t+8)                
                
Title      = ['Passenger vehicles','residential buildings']
ScensL     = ['SSP2, no REFs','SSP2, full REF spectrum','SSP1, no REFs','SSP1, full REF spectrum','LED, no REFs','LED, full REF spectrum']

#mS = 1
#mR = 1
mRCP = 1 # select RCP2.6, which has full implementation of RE strategies by 2050.
for mR in range(0,2): # Veh/Buildings
    
    if mR == 0:
        Data = AnnEmsV_SecondarySteel[:,:,mRCP,:]
    if mR == 1:
        Data = AnnEmsB_SecondarySteel[:,:,mRCP,:]


    fig  = plt.figure(figsize=(8,5))
    ax1  = plt.axes([0.08,0.08,0.85,0.9])
    
    ProxyHandlesList = []   # For legend     
    
    for mS in range(NS-1,-1,-1):
        ax1.plot(np.arange(2016,2051), Data[:,mS,0],  linewidth = linewidth[mS],  linestyle = '-',  color = MyColorCycle[ColorOrder[mS],:])
        #ProxyHandlesList.append(plt.line((0, 0), 1, 1, fc=MyColorCycle[m,:]))
        ax1.plot(np.arange(2016,2051), Data[:,mS,-1], linewidth = linewidth2[mS], linestyle = '--', color = MyColorCycle[ColorOrder[mS],:])
        #ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))     
    plt_lgd  = plt.legend(ScensL,shadow = False, prop={'size':14}, loc = 'upper left',bbox_to_anchor=(1.05, 1))    
    plt.ylabel('Secondary steel and iron, Mt/yr.', fontsize = 18) 
    plt.xlabel('year', fontsize = 18)         
    plt.title('Secondary steel, by socio-economic scenario, \n' + Region + ', ' + Title[mR] + '.', fontsize = 18)
    plt.xticks(fontsize=18)
    plt.yticks(fontsize=18)
    ax1.set_xlim([2015, 2051])
    plt.gca().set_ylim(bottom=0)
    
    plt.show()
    fig_name = 'SecondarySteel_TimeSeries_' + Region + '_ ' + Title[mR] + '.png'
    fig.savefig('C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + fig_name, dpi = 400, bbox_inches='tight')             


