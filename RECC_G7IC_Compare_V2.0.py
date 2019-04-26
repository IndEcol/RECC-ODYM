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
# 2) + EoL + FYI + ReU (+LTE)
# 3) + EoL + FYI + ReU (+LTE) + MSu
# 4) + EoL + FYI + ReU (+LTE) + MSu + LWE
# 5) + EoL + FYI + ReU (+LTE) + MSu + LWE + MIU = ALL 

# LTE for buildings only.

Region= 'Canada'
Scope = 'Canada_Vehicles'
FolderlistV =[
'Canada_2019_4_25__16_5_45',
'Canada_2019_4_25__16_10_23',
'Canada_2019_4_25__16_12_0',
'Canada_2019_4_25__16_13_46',
'Canada_2019_4_25__16_15_1'
]

Scope = 'Canada_Buildings'
FolderlistB =[
'Canada_2019_4_25__16_17_22',
'Canada_2019_4_25__16_19_35',
'Canada_2019_4_25__16_20_48',
'Canada_2019_4_25__16_21_59',
'Canada_2019_4_25__16_23_4'
]


Region= 'USA'
Scope = 'USA_Vehicles'
FolderlistV =[
'USA_2019_4_25__15_26_48',
'USA_2019_4_25__15_29_21',
'USA_2019_4_25__15_30_43',
'USA_2019_4_25__15_31_50',
'USA_2019_4_25__15_32_56'
]

Scope = 'USA_buildings'
FolderlistB =[
'USA_2019_4_25__15_34_4',
'USA_2019_4_25__15_36_1',
'USA_2019_4_25__15_37_10',
'USA_2019_4_25__15_38_21',
'USA_2019_4_25__15_39_34'
]


Region= 'France'
Scope = 'France_Vehicles'
FolderlistV =[
'France_2019_4_25__15_43_17',
'France_2019_4_25__15_53_52',
'France_2019_4_25__15_55_5',
'France_2019_4_25__15_56_30',
'France_2019_4_25__15_57_58'
]

Scope = 'France_Buildings'
FolderlistB =[
'France_2019_4_25__15_59_34',
'France_2019_4_25__16_0_55',
'France_2019_4_25__16_2_12',
'France_2019_4_25__16_3_18',
'France_2019_4_25__16_4_30'
]


Region= 'Germany'
Scope = 'Germany_Vehicles'
FolderlistV =[
'Germany_2019_4_25__16_25_30',
'Germany_2019_4_25__16_29_7',
'Germany_2019_4_25__16_30_11',
'Germany_2019_4_25__16_31_17',
'Germany_2019_4_25__16_32_27'
]

Scope = 'Germany_Buildings'
FolderlistB =[
'Germany_2019_4_25__16_33_36',
'Germany_2019_4_25__16_34_54',
'Germany_2019_4_25__16_35_58',
'Germany_2019_4_25__16_37_7',
'Germany_2019_4_25__16_38_43'
]


Region= 'Japan'
Scope = 'Japan_Vehicles'
FolderlistV =[
'Japan_2019_4_25__17_35_40',
'Japan_2019_4_25__17_38_50',
'Japan_2019_4_25__17_40_28',
'Japan_2019_4_25__17_43_1',
'Japan_2019_4_25__17_45_38'
]

Scope = 'Japan_Buildings'
FolderlistB =[
'Japan_2019_4_25__17_47_14',
'Japan_2019_4_25__17_48_34',
'Japan_2019_4_25__17_50_16',
'Japan_2019_4_25__17_51_37',
'Japan_2019_4_25__17_52_51'
]


Region= 'Italy'
Scope = 'Italy_Vehicles'
FolderlistV =[
'Italy_2019_4_25__17_19_53',
'Italy_2019_4_25__17_21_19',
'Italy_2019_4_25__17_22_32',
'Italy_2019_4_25__17_23_36',
'Italy_2019_4_25__17_24_38'
]

Scope = 'Italy_Buildings'
FolderlistB =[
'Italy_2019_4_25__17_25_51',
'Italy_2019_4_25__17_27_6',
'Italy_2019_4_25__17_28_7',
'Italy_2019_4_25__17_32_51',
'Italy_2019_4_25__17_34_2'
]


Region= 'UK'
Scope = 'UK_Vehicles'
FolderlistV =[
'UK_2019_4_25__17_58_24',
'UK_2019_4_25__18_2_5',
'UK_2019_4_25__18_3_18',
'UK_2019_4_25__18_4_57',
'UK_2019_4_25__18_6_32'
]

Scope = 'UK_Buildings'
FolderlistB =[
'UK_2019_4_25__18_8_4',
'UK_2019_4_25__18_11_4',
'UK_2019_4_25__18_12_20',
'UK_2019_4_25__18_14_50',
'UK_2019_4_25__18_16_10'
]


Region= 'G7'
Scope = 'G7 Vehicles'
FolderlistV =[
'G7_2019_4_25__18_21_20',
'G7_2019_4_25__18_27_44',
'G7_2019_4_25__18_37_49',
'G7_2019_4_25__18_41_35',
'G7_2019_4_25__18_46_2'
]

Scope = 'G7 Buildings'
FolderlistB =[
'G7_2019_4_25__18_49_22',
'G7_2019_4_25__18_52_54',
'G7_2019_4_25__18_56_20',
'G7_2019_4_25__18_59_50',
'G7_2019_4_25__19_3_29'
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
LWE   = ['No RE','higher yields, re-use','material subst.','down-sizing','More intense use','All RE stratgs.']

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
LWE_area   = ['higher yields, re-use','material subst.','down-sizing','more intense use']     

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


