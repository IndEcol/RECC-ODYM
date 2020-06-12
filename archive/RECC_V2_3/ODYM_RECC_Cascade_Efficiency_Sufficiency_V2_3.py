# -*- coding: utf-8 -*-
"""
Created on Wed Oct 17 10:37:00 2018

@author: spauliuk
"""
def main(RegionalScope,ThreeSectoList_Export,SingleSectList):
    
    import xlrd
    import numpy as np
    import matplotlib.pyplot as plt  
    import pylab
    import os
    import RECC_Paths # Import path file   
    
    # FileOrder:
    # 1) None, choose no climate policy scenario
    # 1) None, choose RCP 2.6 scenario
    # 3) Supply-Side ME only
    # 4) Supply and Demand-side ME
    # 5) ALL, including sufficiency
    
    Region      = RegionalScope
    FolderlistV = [SingleSectList[2],ThreeSectoList_Export[0],SingleSectList[0],SingleSectList[3],ThreeSectoList_Export[-1]]
    
    # Waterfall plots.
    
    NS = 3 # no of SSP scenarios
    NR = 2 # no of RCP scenarios
    NE = 5 # no of Res. eff. scenarios
    
    CumEmsV           = np.zeros((NS,NR,NE))   # SSP-Scenario x RCP scenario x RES scenario
    CumEmsV2060       = np.zeros((NS,NR,NE))   # SSP-Scenario x RCP scenario x RES scenario
    AnnEmsV2030       = np.zeros((NS,NR,NE))   # SSP-Scenario x RCP scenario x RES scenario
    AnnEmsV2050       = np.zeros((NS,NR,NE))   # SSP-Scenario x RCP scenario x RES scenario
    ASummaryV         = np.zeros((12,NE))      # For direct copy-paste to Excel
    AvgDecadalEmsV    = np.zeros((NS,NR,NE,4)) # SSP-Scenario x RCP scenario x RES scenario
    
    for r in range(0,NE): # RE scenario
        Path = os.path.join(RECC_Paths.results_path,FolderlistV[r],'SysVar_TotalGHGFootprint.xls')
        Resultfile = xlrd.open_workbook(Path)
        Resultsheet = Resultfile.sheet_by_name('TotalGHGFootprint')
        for s in range(0,NS): # SSP scenario
            for c in range(0,NR):
                for t in range(0,35): # time until 2050 only!!! Cum. emissions until 2050.
                    CumEmsV[s,c,r] += Resultsheet.cell_value(t +2, 1 + c + NR*s)
                for t in range(0,45): # time until 2060.
                    CumEmsV2060[s,c,r] += Resultsheet.cell_value(t +2, 1 + c + NR*s)                    
                AnnEmsV2030[s,c,r]  = Resultsheet.cell_value(16  , 1 + c + NR*s)
                AnnEmsV2050[s,c,r]  = Resultsheet.cell_value(36  , 1 + c + NR*s)
            AvgDecadalEmsV[s,1,r,0]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(7,17)])/10
            AvgDecadalEmsV[s,1,r,1]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(17,27)])/10
            AvgDecadalEmsV[s,1,r,2]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(27,37)])/10
            AvgDecadalEmsV[s,1,r,3]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(37,47)])/10      
            AvgDecadalEmsV[s,0,r,0]   = sum([Resultsheet.cell_value(i, 2*(s+1)-1) for i in range(7,17)])/10
            AvgDecadalEmsV[s,0,r,1]   = sum([Resultsheet.cell_value(i, 2*(s+1)-1) for i in range(17,27)])/10
            AvgDecadalEmsV[s,0,r,2]   = sum([Resultsheet.cell_value(i, 2*(s+1)-1) for i in range(27,37)])/10
            AvgDecadalEmsV[s,0,r,3]   = sum([Resultsheet.cell_value(i, 2*(s+1)-1) for i in range(37,47)])/10               
                
    ASummaryV[0:3,:] = AnnEmsV2030[:,1,:].copy() # RCP is fixed: RCP2.6
    ASummaryV[3:6,:] = AnnEmsV2050[:,1,:].copy() # RCP is fixed: RCP2.6
    ASummaryV[6:9,:] = CumEmsV[:,1,:].copy()     # RCP is fixed: RCP2.6
    ASummaryV[9::,:] = CumEmsV2060[:,1,:].copy() # RCP is fixed: RCP2.6
                        
    # Waterfall plot            
    MyColorCycle = pylab.cm.tab20(np.arange(0,1,0.05)) # select 20 colors from the 'tab20' color map.                       
    
    Sector = ['suff_eff']
    Title  = ['Cum_GHG_2016_2050','Cum_GHG_2040_2050','Annual_GHG_2050']
    Scens  = ['LED','SSP1','SSP2']
    LWE    = ['No climate policy','energy efficiency','energy supply', 'supply-side ME','demand-side ME','sufficiency','residual']
    
    for nn in range(0,3):
        Data = np.zeros((3,6))
        if nn == 0:
            Data[:,0] = CumEmsV[:,0,0].copy()
            Data[:,1] = CumEmsV[:,0,1].copy()
            Data[:,2] = CumEmsV[:,1,1].copy()
            Data[:,3] = CumEmsV[:,1,2].copy()
            Data[:,4] = CumEmsV[:,1,3].copy()
            Data[:,5] = CumEmsV[:,1,4].copy()
        if nn == 1:
            Data[:,0] = 10*AvgDecadalEmsV[:,0,0,2].copy()
            Data[:,1] = 10*AvgDecadalEmsV[:,0,1,2].copy()
            Data[:,2] = 10*AvgDecadalEmsV[:,1,1,2].copy()
            Data[:,3] = 10*AvgDecadalEmsV[:,1,2,2].copy()
            Data[:,4] = 10*AvgDecadalEmsV[:,1,3,2].copy()            
            Data[:,5] = 10*AvgDecadalEmsV[:,1,4,2].copy()            
        if nn == 2:
            Data[:,0] = AnnEmsV2050[:,0,0].copy()
            Data[:,1] = AnnEmsV2050[:,0,1].copy()
            Data[:,2] = AnnEmsV2050[:,1,1].copy()
            Data[:,3] = AnnEmsV2050[:,1,2].copy()
            Data[:,4] = AnnEmsV2050[:,1,3].copy()
            Data[:,5] = AnnEmsV2050[:,1,4].copy()
            
        Left  = Data[2,0]
        
        Xoffs1 = [1,3,5]
        Xoffs2 = [1.7,3.7,5.7]
        
        bw = 0.5
        
        fig  = plt.figure(figsize=(5,8))
        ax1  = plt.axes([0.08,0.08,0.85,0.9])
        
        for ms in range (0,NS):
            # plot bars
            ax1.fill_between([Xoffs1[ms],Xoffs1[ms]+bw], [0,0],[Data[ms,0],Data[ms,0]],linestyle = '--', facecolor =MyColorCycle[11,:], linewidth = 0.0)
            ax1.fill_between([Xoffs2[ms],Xoffs2[ms]+bw], [Data[ms,1],Data[ms,1]],[Data[ms,0],Data[ms,0]],linestyle = '--', facecolor =MyColorCycle[4,:], linewidth = 0.0)
            ax1.fill_between([Xoffs2[ms],Xoffs2[ms]+bw], [Data[ms,2],Data[ms,2]],[Data[ms,1],Data[ms,1]],linestyle = '--', facecolor =MyColorCycle[0,:], linewidth = 0.0)
            ax1.fill_between([Xoffs2[ms],Xoffs2[ms]+bw], [Data[ms,3],Data[ms,3]],[Data[ms,2],Data[ms,2]],linestyle = '--', facecolor =MyColorCycle[2,:], linewidth = 0.0)
            ax1.fill_between([Xoffs2[ms],Xoffs2[ms]+bw], [Data[ms,4],Data[ms,4]],[Data[ms,3],Data[ms,3]],linestyle = '--', facecolor =MyColorCycle[3,:], linewidth = 0.0)
            ax1.fill_between([Xoffs2[ms],Xoffs2[ms]+bw], [Data[ms,5],Data[ms,5]],[Data[ms,4],Data[ms,4]],linestyle = '--', facecolor =MyColorCycle[6,:], linewidth = 0.0)
            ax1.fill_between([Xoffs2[ms],Xoffs2[ms]+bw], [0,0],[Data[ms,5],Data[ms,5]],linestyle = '--', facecolor =MyColorCycle[15,:], linewidth = 0.0)
            if ms == 1: 
                ProxyHandlesList = []   # For legend     
                ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[11,:])) # create proxy artist for legend
                ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[4,:])) # create proxy artist for legend
                ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[0,:])) # create proxy artist for legend
                ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[2,:])) # create proxy artist for legend
                ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[3,:])) # create proxy artist for legend
                ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[6,:])) # create proxy artist for legend
                ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[15,:])) # create proxy artist for legend
        

    
#            # plot lines:
#            plt.plot([0,8.5],[Left,Left],linestyle = '-', linewidth = 0.5, color = 'k')
#            plt.plot([1,2.5],[Data[1,m],Data[1,m]],linestyle = '-', linewidth = 0.5, color = 'k')
#            plt.plot([2,3.5],[Data[2,m],Data[2,m]],linestyle = '-', linewidth = 0.5, color = 'k')
#            plt.plot([3,4.5],[Data[3,m],Data[3,m]],linestyle = '-', linewidth = 0.5, color = 'k')
#            plt.plot([4,5.5],[Data[4,m],Data[4,m]],linestyle = '-', linewidth = 0.5, color = 'k')
#            plt.plot([5,6.5],[Data[5,m],Data[5,m]],linestyle = '-', linewidth = 0.5, color = 'k')
#            plt.plot([6,7.5],[Data[6,m],Data[6,m]],linestyle = '-', linewidth = 0.5, color = 'k')
#            plt.plot([7,8.5],[Data[7,m],Data[7,m]],linestyle = '-', linewidth = 0.5, color = 'k')
#    
#            plt.arrow(8.25, Data[7,m],0, Data[0,m]-Data[7,m], lw = 0.8, ls = '-', shape = 'full',
#                  length_includes_head = True, head_width =0.1, head_length =0.01*Left, ec = 'k', fc = 'k')
#            plt.arrow(8.25,Data[0,m],0,Data[7,m]-Data[0,m], lw = 0.8, ls = '-', shape = 'full',
#                  length_includes_head = True, head_width =0.1, head_length =0.01*Left, ec = 'k', fc = 'k')

        # plot text and labels
        #plt.text(6.85, 0.94 *Left, ("%3.0f" % inc) + ' %',fontsize=18,fontweight='bold')          
        #plt.text(4.3, 0.94  *Right, Scens[m],fontsize=18,fontweight='bold') 
        #plt.title('Energy, efficiency, and sufficiency, ' + Sector[0] + '.', fontsize = 18)
        plt.ylabel(Title[nn] + ', Mt.', fontsize = 18)
        plt.xticks([1.6,3.6,5.6])
        plt.yticks(fontsize =18)
        ax1.set_xticklabels(Scens, rotation =0, fontsize = 21, fontweight = 'normal')
        plt_lgd  = plt.legend(handles = ProxyHandlesList,labels = LWE,shadow = False, prop={'size':12},ncol=1, loc = 'upper left' ) 
        #plt.axis([-0.2, 7.7, 0.9*Right, 1.02*Left])
        plt.axis([-0.2, 7, 0, 1.03*Left])
    
        plt.show()
        fig_name = Title[nn] + Region + '_ ' + Sector[0] + '.png'
        fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = 400, bbox_inches='tight')             
        
    
 
    
    return None

# code for script to be run as standalone function
if __name__ == "__main__":
    main()
