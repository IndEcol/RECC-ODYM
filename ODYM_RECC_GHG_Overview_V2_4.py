# -*- coding: utf-8 -*-
"""
Created on Thu Sep  3 11:18:10 2020

@author: spauliuk
"""

"""
File ODYM_RECC_GHG_Overview_V2_4.py

Script that runs the sensitivity and scnenario comparison scripts for different settings.


"""

# Import required libraries:
import os
import numpy as np
import matplotlib.pyplot as plt 
import pylab

import RECC_Paths # Import path file

# plot GHG overview figure global
def main(RegionalScope,SectorString,CumEms2050,CumEms2060,TimeSeries_R,PlotExpResolution,NE,LWE_Labels,Current_UUID):
    RECC_Paths.results_path_save = os.path.join(RECC_Paths.results_path_eval,'RECC_Results_' + Current_UUID)
    # Color def:
    BaseBlue      = np.array([0.208,0.592,0.561,1]) # Base for GHG after full ME reduction
    #BaseRed       = np.array([0.48,0.33,0.22,1])
    ColOrder     = [11,4,0,18,8,16,2,6,15] # for all sector selection other than pav and reb.
    LabelColors  = ['k','k','k','k','k','k','k',BaseBlue]
    if SectorString == 'pav':
        ColOrder     = [11,0,18,8,16,2,6,15]
        LabelColors  = ['k','k','k','k','k','k',BaseBlue]
    if SectorString == 'reb':
        ColOrder     = [11,4,8,16,2,6,15]
        LabelColors  = ['k','k','k','k','k',BaseBlue]
    TSList           = ['2050','2060']
    for TS in range(0,2): # TS: temporal scope: 0: 2050, 1: 2060
        Xoff         = [1,1.7,2.7,3.4,4.4,5.1]
        #MyColorCycle = pylab.cm.tab20(np.arange(0,1,0.05)) # select 20 colors from the 'tab20' color map. 
        MyColorCycle = np.zeros((20,4))
        # Define Colors:
        MyColorCycle[6,:]  = np.array([0.84313725, 0.188235294,0.152941176,1]) # See https://colorbrewer2.org/#type=diverging&scheme=RdYlBu&n=7
        MyColorCycle[2,:]  = np.array([0.988235294,0.552941176,0.349019608,1])
        MyColorCycle[16,:] = np.array([0.996078431,0.878431373,0.564705882,1])
        MyColorCycle[8,:]  = np.array([1,          1,          0.749019608,1])
        MyColorCycle[18,:] = np.array([0.878431373,0.952941176,0.97254902,1])
        MyColorCycle[0,:]  = np.array([0.568627451,0.749019608,0.858823529,1])
        MyColorCycle[4,:]  = np.array([0.270588235,0.458823529,0.705882353,1])
        MyColorCycle[11,:] = np.array([0.8,0.8,0.8,1]) # grey
        if TS == 0:
            Data_Cum_Abs = CumEms2050
            Data_Cum_pc  = Data_Cum_Abs / np.einsum('SR,E->SRE',Data_Cum_Abs[:,:,0],np.ones(NE)) # here: pc = percent, not per capita.
            Data_50_Abs  = np.einsum('ESR->SRE',TimeSeries_R[0,:,34,:,:]) # starts at 0 for 2016
            Data_50_pc   = Data_50_Abs / np.einsum('SR,E->SRE',Data_50_Abs[:,:,0],np.ones(NE))   # here: pc = percent, not per capita.
        if TS == 1:
            Data_Cum_Abs = CumEms2060
            Data_Cum_pc  = Data_Cum_Abs / np.einsum('SR,E->SRE',Data_Cum_Abs[:,:,0],np.ones(NE)) # here: pc = percent, not per capita.
            Data_50_Abs  = np.einsum('ESR->SRE',TimeSeries_R[0,:,34,:,:]) # starts at 0 for 2016
            Data_50_pc   = Data_50_Abs / np.einsum('SR,E->SRE',Data_50_Abs[:,:,0],np.ones(NE))   # here: pc = percent, not per capita.
        Cum_savings  = Data_Cum_Abs[:,:,0] - Data_Cum_Abs[:,:,-1]
        Ann_savings  = Data_50_Abs[:,:,0]  - Data_50_Abs[:,:,-1]
        LWE          = LWE_Labels
        XTicks       = [1.25,1.95,2.95,3.65,4.65,5.35]
        YTicks       = [0,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9,1.0,1.2,1.3,1.4,1.5,1.6,1.7,1.8,1.9,2.0,2.1,2.2]
        YTickLabels  = ['0','10','20','30','40','50','60','70','80','90','100','0','10','20','30','40','50','60','70','80','90','100']
        # plot results 
        bw   = 0.5
        
        yoff = 1.2
        fig  = plt.figure(figsize=(8,5))
        ax1  = plt.axes([0.08,0.08,0.85,0.9])
    
        ProxyHandlesList = []   # For legend     
        # plot bars
        for mS in range(0,3):
            for mR in range(0,2):
                ax1.fill_between([Xoff[mS*2+mR],Xoff[mS*2+mR]+bw], [yoff,yoff],[yoff+Data_Cum_pc[mS,mR,-1],yoff+Data_Cum_pc[mS,mR,-1]],linestyle = '--', facecolor =MyColorCycle[ColOrder[0],:], linewidth = 0.0)
                for xca in range(1,NE):
                    ax1.fill_between([Xoff[mS*2+mR],Xoff[mS*2+mR]+bw], [yoff+Data_Cum_pc[mS,mR,-xca],yoff+Data_Cum_pc[mS,mR,-xca]],[yoff+Data_Cum_pc[mS,mR,-xca-1],yoff+Data_Cum_pc[mS,mR,-xca-1]],linestyle = '--', facecolor =MyColorCycle[ColOrder[xca],:], linewidth = 0.0)
                        
                ax1.fill_between([Xoff[mS*2+mR],Xoff[mS*2+mR]+bw], [0,0],[Data_50_pc[mS,mR,-1],Data_50_pc[mS,mR,-1]],linestyle = '--', facecolor =MyColorCycle[ColOrder[0],:], linewidth = 0.0)
                for xca in range(1,NE):
                    ax1.fill_between([Xoff[mS*2+mR],Xoff[mS*2+mR]+bw], [Data_50_pc[mS,mR,-xca],Data_50_pc[mS,mR,-xca]],[Data_50_pc[mS,mR,-xca-1],Data_50_pc[mS,mR,-xca-1]],linestyle = '--', facecolor =MyColorCycle[ColOrder[xca],:], linewidth = 0.0)
                
        plt.plot([-1,8],[1.1,1.1],  linestyle = '-',  linewidth = 0.5, color = 'k')
        plt.plot([2.45,2.45],[-1,3],linestyle = '--', linewidth = 0.5, color = 'k')
        plt.plot([4.15,4.15],[-1,3],linestyle = '--', linewidth = 0.5, color = 'k')
        for fca in range(0,NE):
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[ColOrder[fca],:])) # create proxy artist for legend
            #ProxyHandlesList.append(plt.plot([], [], ' '))
        # plot emissions reductions
        for mS in range(0,3):
            for mR in range(0,2): # 1) Cum. GHG savings, Gt, top; 2) Ann. GHG savings, Mt, bottom; 3) Residual Cum. GHG, Gt, top; 4) Residual Ann. GHG, Mt, bottom;
                plt.text(Xoff[mS*2+mR]+0.51, 2.19,  ("%3.0f" % (Cum_savings[mS,mR]/1000))                   + ' Gt',fontsize=14, rotation=90, fontweight='bold')
                plt.text(Xoff[mS*2+mR]+0.51, 0.99,  ("%3.0f" % (10*np.round(Ann_savings[mS,mR]/10)))        + ' Mt',fontsize=14, rotation=90, fontweight='bold')
                plt.text(Xoff[mS*2+mR]+0.2, 1.68,   ("%3.0f" % (10*np.round(Data_Cum_Abs[mS,mR,-1]/10000))) + ' Gt',fontsize=16, rotation=90, fontweight='normal', color = BaseBlue)
                if mS*2+mR == 1: # fine tune for paper plot, LED/2°C and SSP2/NoPol are adjusted separately
                    plt.text(Xoff[mS*2+mR]+0.18, 0.17+0.05,   ("%3.0f" % (100*np.round(Data_50_Abs[mS,mR,-1]/100)))          ,fontsize=15, rotation=90, fontweight='normal', color = BaseBlue)
                    plt.text(Xoff[mS*2+mR]+0.32, 0.17-0.01,   'Mt'                                                           ,fontsize=15, rotation=90, fontweight='normal', color = BaseBlue)
                elif mS*2+mR == 4:
                    plt.text(Xoff[mS*2+mR]+0.2, 0.52+0.02,   ("%3.0f" % (100*np.round(Data_50_Abs[mS,mR,-1]/100)))   + ' Mt',fontsize=15, rotation=90, fontweight='normal', color = BaseBlue)
                else:
                    plt.text(Xoff[mS*2+mR]+0.2, 0.45+0.02,   ("%3.0f" % (100*np.round(Data_50_Abs[mS,mR,-1]/100)))   + ' Mt',fontsize=15, rotation=90, fontweight='normal', color = BaseBlue)
                    
                # Scenarios
                plt.text(1.4, -0.15,     'LED', fontsize=17, fontweight='bold', color = 'k')
                plt.text(3.05, -0.15,    'SSP1',fontsize=17, fontweight='bold', color = 'k')
                plt.text(4.75, -0.15,    'SSP2',fontsize=17, fontweight='bold', color = 'k')
                # top arrow, black
                plt.arrow(Xoff[mS*2+mR]+0.46,2.25,0,-0.05, lw = 0.8, ls = '-', shape = 'full',
                  length_includes_head = True, head_width =0.06, head_length =0.02, ec = 'k', fc = 'k')
                plt.arrow(Xoff[mS*2+mR]+0.46,1.05,0,-0.05, lw = 0.8, ls = '-', shape = 'full',
                  length_includes_head = True, head_width =0.06, head_length =0.02, ec = 'k', fc = 'k')
                # bottom arrow, black
                plt.arrow(Xoff[mS*2+mR]+0.46,1.2+Data_Cum_pc[mS,mR,-1]-0.05,0,0.05, lw = 0.8, ls = '-', shape = 'full',
                  length_includes_head = True, head_width =0.06, head_length =0.02, ec = 'k', fc = 'k')
                plt.arrow(Xoff[mS*2+mR]+0.46,Data_50_pc[mS,mR,-1]-0.05,0,0.05, lw = 0.8, ls = '-', shape = 'full',
                  length_includes_head = True, head_width =0.06, head_length =0.02, ec = 'k', fc = 'k')
                
#                plt.arrow(Xoff[mS*2+mR]+0.25,1.2+Data_Cum_pc[mS,mR,-1]-0.05,0,0.05, lw = 0.8, ls = '-', shape = 'full',
#                  length_includes_head = True, head_width =0.06, head_length =0.02, ec = BaseBlue, fc = BaseBlue)
#                plt.arrow(Xoff[mS*2+mR]+0.25,Data_50_pc[mS,mR,-1]-0.05,0,0.05, lw = 0.8, ls = '-', shape = 'full',
#                  length_includes_head = True, head_width =0.06, head_length =0.02, ec = BaseBlue, fc = BaseBlue)
                # top arrow, blue
                plt.arrow(Xoff[mS*2+mR]+0.25,1.2+Data_Cum_pc[mS,mR,-1]-0.05,0,0.05, lw = 0.8, ls = '-', shape = 'full',
                  length_includes_head = True, head_width =0.06, head_length =0.02, ec = BaseBlue, fc = BaseBlue)
                plt.arrow(Xoff[mS*2+mR]+0.25,Data_50_pc[mS,mR,-1]-0.05,0,0.05, lw = 0.8, ls = '-', shape = 'full',
                  length_includes_head = True, head_width =0.06, head_length =0.02, ec = BaseBlue, fc = BaseBlue)
                # bottom arrow, blue
                plt.arrow(Xoff[mS*2+mR]+0.25,1.2+0.05,0,-0.05, lw = 0.8, ls = '-', shape = 'full',
                  length_includes_head = True, head_width =0.06, head_length =0.02, ec = BaseBlue, fc = BaseBlue)
                plt.arrow(Xoff[mS*2+mR]+0.25,0.05,0,-0.05, lw = 0.8, ls = '-', shape = 'full',
                  length_includes_head = True, head_width =0.06, head_length =0.02, ec = BaseBlue, fc = BaseBlue)
                
                #plt.text(Xoff[mS*2+mR]+0.2, 1.9, ("%3.0f" % (Cum_savings[mS,mR]/1000)) + ' Gt',fontsize=16, rotation=90, fontweight='normal')
                #plt.text(Xoff[mS*2+mR]+0.2, 0.5, ("%3.0f" % Ann_savings[mS,mR]) + ' Mt',fontsize=16, rotation=90, fontweight='normal')
                
        # plot text and labels
        plt.text(0.85, 0.97,     '2050 annual emissions', fontsize=11, rotation=90, fontweight='bold') 
        if TS == 0:         
            plt.text(0.85, 2.2,  '2016-50 cum. emissions',fontsize=11, rotation=90, fontweight='bold')          
        if TS == 1:
            plt.text(0.85, 2.2,  '2016-60 cum. emissions',fontsize=11, rotation=90, fontweight='bold')          
    
        #plt.title('RE strats. and GHG emissions, global, pav+reb.', fontsize = 18)
        
        plt.ylabel('GHG emissions, system-wide, %.', fontsize = 18)
        plt.xticks(XTicks)
        plt.yticks(YTicks, fontsize =12)
        #ax1.set_xticklabels(['LED/NoPol','LED/RCP2.6','SSP1/NoPol','SSP1/RCP2.6','SSP2/NoPol','SSP2/RCP2.6'], rotation =90, fontsize = 15, fontweight = 'normal')
        ax1.set_xticklabels(['No Pol.','2°C Pol.','No Pol.','2°C Pol.','No Pol.','2°C Pol.'], rotation =0, fontsize = 14, fontweight = 'bold')
        ax1.set_yticklabels(YTickLabels, rotation =0, fontsize = 15, fontweight = 'normal')
        leg = plt.legend(handles = reversed(ProxyHandlesList),labels = LWE,shadow = False, prop={'size':12}, ncol=1, loc = 'upper right' , bbox_to_anchor=(1.40, 1)) 
        # Change text color:
        mc = 0
        for text in leg.get_texts():
            plt.setp(text, color = LabelColors[mc])
            mc +=1
        #plt.axis([-0.2, 7.7, 0.9*Right, 1.02*Left])
        plt.axis([0.7, 5.8, -0.2, 2.3])
    
        plt.show()
        fig_name = RegionalScope + '_' + SectorString + '_GHG_Overview_rel_' + TSList[TS] + '.png'
        fig.savefig(os.path.join(RECC_Paths.results_path_save,fig_name), dpi = PlotExpResolution, bbox_inches='tight')   
        fig_name = RegionalScope + '_' + SectorString + '_GHG_Overview_rel_' + TSList[TS] + '.svg'
        fig.savefig(os.path.join(RECC_Paths.results_path_save,fig_name), dpi = PlotExpResolution, bbox_inches='tight')   


#
#
#
#
#
#
#
