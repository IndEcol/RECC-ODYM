# -*- coding: utf-8 -*-
"""
Created on Wed Oct 17 10:37:00 2018

@author: spauliuk
"""

import xlrd
import numpy as np
import matplotlib.pyplot as plt  
import pylab

# EU
#FolderlistV =[
#        'G7IC_V1_2019_2_14__10_12_35',
#        'G7IC_V1_2019_2_14__10_19_30',
#        'G7IC_V1_2019_2_14__10_26_36',
#        'G7IC_V1_2019_2_14__10_33_45',
#        'G7IC_V1_2019_2_14__9_51_25'
#        ]

# USA
FolderlistV =[
'G7IC_V1_2019_2_14__10_51_47',
'G7IC_V1_2019_2_14__11_0_57',
'G7IC_V1_2019_2_14__11_9_41',
'G7IC_V1_2019_2_14__11_14_1',
'G7IC_V1_2019_2_14__9_48_27'
]

# 1) None
# 2) + IU
# 3) + IU + LWE
# 4) + IU + LWE + LTE
# 5) + IU + LWE + LTE + EoL-RR = ALL
NS = 3
NR = 5

CumEmsV = np.zeros((NR,NS)) # RE-Scenario x SSP scenario

for m in range(0,NR): # RE scenario
    Path = 'C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + FolderlistV[m] + '\\'
    Resultfile = xlrd.open_workbook(Path + 'SysVar_TotalGHGFootprint.xls')
    Resultsheet = Resultfile.sheet_by_name('TotalGHGFootprint')
    for n in range(0,NS): # SSP scenario
        for o in range(0,35): # time
            CumEmsV[m,n] += Resultsheet.cell_value(o +2, NS*n +1)
    
    
AnnEmsV2050 = np.zeros((NR,NS)) # RE-Scenario x SSP scenario

for m in range(0,NR): # RE scenario
    Path = 'C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + FolderlistV[m] + '\\'
    Resultfile = xlrd.open_workbook(Path + 'SysVar_TotalGHGFootprint.xls')
    Resultsheet = Resultfile.sheet_by_name('TotalGHGFootprint')
    for n in range(0,NS): # SSP scenario
        AnnEmsV2050[m,n] = Resultsheet.cell_value(36, NS*n +1)    

FolderlistB =FolderlistV

# 1) None
# 2) + IU + RU
# 3) + IU + RU + LWE
# 4) + IU + RU + LWE + LTE
# 5) + IU + RU + LWE + LTE + EoL-RR = ALL

CumEmsB = np.zeros((NR,NS)) # RE-Scenario x SSP scenario

for m in range(0,NR): # RE scenario
    Path = 'C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + FolderlistB[m] + '\\'
    Resultfile = xlrd.open_workbook(Path + 'SysVar_TotalGHGFootprint.xls')
    Resultsheet = Resultfile.sheet_by_name('TotalGHGFootprint')
    for n in range(0,NS): # SSP scenario
        for o in range(0,35): # time
            CumEmsB[m,n] += Resultsheet.cell_value(o +2, NS*n +1)
      
        
AnnEmsB2050 = np.zeros((NR,NS)) # RE-Scenario x SSP scenario

for m in range(0,NR): # RE scenario
    Path = 'C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + FolderlistB[m] + '\\'
    Resultfile = xlrd.open_workbook(Path + 'SysVar_TotalGHGFootprint.xls')
    Resultsheet = Resultfile.sheet_by_name('TotalGHGFootprint')
    for n in range(0,NS): # SSP scenario
        AnnEmsB2050[m,n] = Resultsheet.cell_value(36, NS*n +1)        
            
# Plot            
MyColorCycle = pylab.cm.Set1(np.arange(0,1,0.1)) # select 12 colors from the 'Paired' color map.            

Title = ['All','Passenger vehicles','residential buildings']
Scens = ['SSP1','SSP2','SSP3','SSP4','SSP5']
LWE   = ['No RE','More intense use','light-weighting','lifetime ext.','recycling improvemt.','All RE stratgs.']

# Cumulative emissions
for m in range(0,NS): # SSP
    for n in range(0,1): # All/Veh/Buldings
        
        if n == 0:
            Data = CumEmsV
        if n == 1:
            Data = CumEmsB
    
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
        fig_name = 'Cum_GHG_G7_' + Title[n] + '_' + Scens[m] + '.png'
        fig.savefig('C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + fig_name, dpi = 400, bbox_inches='tight')             
            



# 2050 emissions
for m in range(0,NS): # SSP
    for n in range(0,1): # All/Veh/Buldings
        
        if n == 0:
            Data = AnnEmsV2050
        if n == 1:
            Data = AnnEmsB2050
    
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
        fig_name = 'GHG_2050_G7_' + Title[n] + '_' + Scens[m] + '.png'
        fig.savefig('C:\\Users\\spauliuk\\FILES\\ARBEIT\\PROJECTS\\ODYM-RECC\\RECC_Results\\' + fig_name, dpi = 400, bbox_inches='tight')             







#
#