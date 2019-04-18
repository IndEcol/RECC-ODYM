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
        
        
     


     
# Plot            
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







#
#