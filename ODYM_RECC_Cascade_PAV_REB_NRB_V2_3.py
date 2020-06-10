# -*- coding: utf-8 -*-
"""
Created on Wed Oct 17 10:37:00 2018

@author: spauliuk
"""
def main(RegionalScope,ThreeSectoList):
    
    import xlrd
    import numpy as np
    import matplotlib.pyplot as plt  
    import pylab
    import os
    import RECC_Paths # Import path file   
    
    # FileOrder:
    # 1) None
    # 2) + EoL + FSD + FYI
    # 3) + EoL + FSD + FYI + ReU +LTE
    # 4) + EoL + FSD + FYI + ReU +LTE + MSu
    # 5) + EoL + FSD + FYI + ReU +LTE + MSu + LWE
    # 6) + EoL + FSD + FYI + ReU +LTE + MSu + LWE + CaS 
    # 7) + EoL + FSD + FYI + ReU +LTE + MSu + LWE + CaS + RiS
    # 8) + EoL + FSD + FYI + ReU +LTE + MSu + LWE + CaS + RiS + MIU = ALL 
    
    Region      = RegionalScope
    FolderlistV = ThreeSectoList
    
    # Waterfall plots.
    
    NS = 3 # no of SSP scenarios
    NR = 2 # no of RCP scenarios
    NE = 8 # no of Res. eff. scenarios
    
    CumEmsV           = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    CumEmsV2060       = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    AnnEmsV2030       = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    AnnEmsV2050       = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    ASummaryV         = np.zeros((12,NE)) # For direct copy-paste to Excel
    AvgDecadalEmsV    = np.zeros((NS,NE,4)) # SSP-Scenario x RES scenario, RCP is fixed: RCP2.6
    # for materials:
    MatCumEmsV        = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    MatCumEmsV2060    = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    MatAnnEmsV2030    = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    MatAnnEmsV2050    = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario    
    MatSummaryV       = np.zeros((12,NE)) # For direct copy-paste to Excel
    AvgDecadalMatEmsV = np.zeros((NS,NE,4)) # SSP-Scenario x RES scenario, RCP is fixed: RCP2.6
    # for materials incl. recycling credit:
    MatCumEmsVC       = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    MatCumEmsVC2060   = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    MatAnnEmsV2030C   = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    MatAnnEmsV2050C   = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario    
    MatSummaryVC      = np.zeros((12,NE)) # For direct copy-paste to Excel
    AvgDecadalMatEmsVC= np.zeros((NS,NE,4)) # SSP-Scenario x RES scenario, RCP is fixed: RCP2.6
     
    
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
            AvgDecadalEmsV[s,r,0]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(7,17)])/10
            AvgDecadalEmsV[s,r,1]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(17,27)])/10
            AvgDecadalEmsV[s,r,2]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(27,37)])/10
            AvgDecadalEmsV[s,r,3]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(37,47)])/10                    
                
    ASummaryV[0:3,:] = AnnEmsV2030[:,1,:].copy() # RCP is fixed: RCP2.6
    ASummaryV[3:6,:] = AnnEmsV2050[:,1,:].copy() # RCP is fixed: RCP2.6
    ASummaryV[6:9,:] = CumEmsV[:,1,:].copy()     # RCP is fixed: RCP2.6
    ASummaryV[9::,:] = CumEmsV2060[:,1,:].copy() # RCP is fixed: RCP2.6   
    
    for r in range(0,NE): # RE scenario
        Path = os.path.join(RECC_Paths.results_path,FolderlistV[r],'SysVar_TotalGHGFootprint.xls')
        Resultfile   = xlrd.open_workbook(Path)
        Resultsheet  = Resultfile.sheet_by_name('TotalGHGFootprint')
        Resultsheet1 = Resultfile.sheet_by_name('Cover')
        UUID         = Resultsheet1.cell_value(3,2)
        Resultfile2  = xlrd.open_workbook(os.path.join(RECC_Paths.results_path,FolderlistV[r],'ODYM_RECC_ModelResults_' + UUID + '.xlsx'))
        Resultsheet2 = Resultfile2.sheet_by_name('Model_Results')
        # Find the index for the recycling credit and others:
        rci = 1
        while True:
            if Resultsheet2.cell_value(rci, 0) == 'GHG emissions, recycling credits':
                break # that gives us the right index to read the recycling credit from the result table.
            rci += 1
        mci = 1
        while True:
            if Resultsheet2.cell_value(mci, 0) == 'GHG emissions, material cycle industries and their energy supply _3di_9di':
                break # that gives us the right index from the result table.
            mci += 1
            
        # Material results export
        for s in range(0,NS): # SSP scenario
            for c in range(0,NR):
                for t in range(0,35): # time until 2050 only!!! Cum. emissions until 2050.
                    MatCumEmsV[s,c,r] += Resultsheet2.cell_value(mci+ 2*s +c,t+8)
                for t in range(0,45): # time until 2060.
                    MatCumEmsV2060[s,c,r] += Resultsheet2.cell_value(mci+ 2*s +c,t+8)                    
                MatAnnEmsV2030[s,c,r]  = Resultsheet2.cell_value(mci+ 2*s +c,22)
                MatAnnEmsV2050[s,c,r]  = Resultsheet2.cell_value(mci+ 2*s +c,42)
            AvgDecadalMatEmsV[s,r,0]   = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for t in range(13,23)])/10
            AvgDecadalMatEmsV[s,r,1]   = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for t in range(23,33)])/10
            AvgDecadalMatEmsV[s,r,2]   = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for t in range(33,43)])/10
            AvgDecadalMatEmsV[s,r,3]   = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for t in range(43,53)])/10        
                        
    # Waterfall plot, system-wide emissions   
    MyColorCycle = pylab.cm.tab20(np.arange(0,1,0.05)) # select 20 colors from the 'tab20' color map.            
    
    Sector = ['pav_reb_nrb']
    Title  = ['Cum_GHG_2016_2050','Cum_GHG_2040_2050','Annual_GHG_2050']
    Scens  = ['LED','SSP1','SSP2']
    LWE    = ['No RE','higher yields', 're-use/longer use','material subst.','down-sizing','car-sharing','ride-sharing','More intense bld. use','All RE stratgs.']
    
    for nn in range(0,3):
        for m in range(0,NS): # SSP
            if nn == 0:
                Data = np.einsum('SE->ES',CumEmsV[:,1,:])
            if nn == 1:
                Data = np.einsum('SE->ES',10*AvgDecadalEmsV[:,:,2])
            if nn == 2:
                Data = np.einsum('SE->ES',AnnEmsV2050[:,1,:])
                
            inc = -100 * (Data[0,m] - Data[7,m])/Data[0,m]
        
            Left  = Data[0,m]
            Right = Data[7,m]
            # plot results
            bw = 0.5
            ga = 0.3
        
            fig  = plt.figure(figsize=(5,8))
            ax1  = plt.axes([0.08,0.08,0.85,0.9])
        
            ProxyHandlesList = []   # For legend     
            # plot bars
            ax1.fill_between([0,0+bw], [0,0],[Left,Left],linestyle = '--', facecolor =MyColorCycle[11,:], linewidth = 0.0)
            ax1.fill_between([1,1+bw], [Data[1,m],Data[1,m]],[Left,Left],linestyle = '--', facecolor =MyColorCycle[4,:], linewidth = 0.0)
            ax1.fill_between([2,2+bw], [Data[2,m],Data[2,m]],[Data[1,m],Data[1,m]],linestyle = '--', facecolor =MyColorCycle[0,:], linewidth = 0.0)
            ax1.fill_between([3,3+bw], [Data[3,m],Data[3,m]],[Data[2,m],Data[2,m]],linestyle = '--', facecolor =MyColorCycle[18,:], linewidth = 0.0)
            ax1.fill_between([4,4+bw], [Data[4,m],Data[4,m]],[Data[3,m],Data[3,m]],linestyle = '--', facecolor =MyColorCycle[8,:], linewidth = 0.0)
            ax1.fill_between([5,5+bw], [Data[5,m],Data[5,m]],[Data[4,m],Data[4,m]],linestyle = '--', facecolor =MyColorCycle[16,:], linewidth = 0.0)
            ax1.fill_between([6,6+bw], [Data[6,m],Data[6,m]],[Data[5,m],Data[5,m]],linestyle = '--', facecolor =MyColorCycle[2,:], linewidth = 0.0)
            ax1.fill_between([7,7+bw], [Data[7,m],Data[7,m]],[Data[6,m],Data[6,m]],linestyle = '--', facecolor =MyColorCycle[6,:], linewidth = 0.0)
            ax1.fill_between([8,8+bw], [0,0],[Data[7,m],Data[7,m]],linestyle = '--', facecolor =MyColorCycle[15,:], linewidth = 0.0)
            
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[11,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[4,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[0,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[18,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[8,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[16,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[2,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[6,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[15,:])) # create proxy artist for legend
            
            # plot lines:
            plt.plot([0,8.5],[Left,Left],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([1,2.5],[Data[1,m],Data[1,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([2,3.5],[Data[2,m],Data[2,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([3,4.5],[Data[3,m],Data[3,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([4,5.5],[Data[4,m],Data[4,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([5,6.5],[Data[5,m],Data[5,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([6,7.5],[Data[6,m],Data[6,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([7,8.5],[Data[7,m],Data[7,m]],linestyle = '-', linewidth = 0.5, color = 'k')
    
            plt.arrow(8.25, Data[7,m],0, Data[0,m]-Data[7,m], lw = 0.8, ls = '-', shape = 'full',
                  length_includes_head = True, head_width =0.1, head_length =0.01*Left, ec = 'k', fc = 'k')
            plt.arrow(8.25,Data[0,m],0,Data[7,m]-Data[0,m], lw = 0.8, ls = '-', shape = 'full',
                  length_includes_head = True, head_width =0.1, head_length =0.01*Left, ec = 'k', fc = 'k')
    
            # plot text and labels
            plt.text(6.85, 0.94 *Left, ("%3.0f" % inc) + ' %',fontsize=18,fontweight='bold')          
            plt.text(4.3, 0.94  *Right, Scens[m],fontsize=18,fontweight='bold') 
            plt.title('RE strategies and GHG emissions, ' + Sector[0] + '.', fontsize = 18)
            plt.ylabel(Title[nn] + ', Mt.', fontsize = 18)
            plt.xticks([0.25,1.25,2.25,3.25,4.25,5.25,6.25,7.25,8.25])
            plt.yticks(fontsize =18)
            ax1.set_xticklabels([], rotation =90, fontsize = 21, fontweight = 'normal')
            plt_lgd  = plt.legend(handles = ProxyHandlesList,labels = LWE,shadow = False, prop={'size':12},ncol=1, loc = 'upper right' ,bbox_to_anchor=(1.91, 1)) 
            #plt.axis([-0.2, 7.7, 0.9*Right, 1.02*Left])
            plt.axis([-0.2, 8.7, 0, 1.02*Left])
        
            plt.show()
            fig_name = Title[nn] + Region + '_ ' + Sector[0] + '_' + Scens[m] + '.png'
            fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = 400, bbox_inches='tight')             
                
    
    # Waterfall plot, material cycle emissions   
    MyColorCycle = pylab.cm.tab20(np.arange(0,1,0.05)) # select 12 colors from the 'tab20' color map.            
    
    Sector = ['pav_reb_nrb_MC']
    Title  = ['Cum_GHG_2016_2050_Mat','Cum_GHG_2040_2050_Mat','Annual_GHG_2050_Mat']
    Scens  = ['LED','SSP1','SSP2']
    LWE    = ['No RE','higher yields', 're-use/longer use','material subst.','down-sizing','car-sharing','ride-sharing','More intense bld. use','All RE stratgs.']
    
    for nn in range(0,3):
        for m in range(0,NS): # SSP
            if nn == 0:
                Data = np.einsum('SE->ES',MatCumEmsV[:,1,:])
            if nn == 1:
                Data = np.einsum('SE->ES',10*AvgDecadalMatEmsV[:,:,2])
            if nn == 2:
                Data = np.einsum('SE->ES',MatAnnEmsV2050[:,1,:])
                
            inc = -100 * (Data[0,m] - Data[7,m])/Data[0,m]
        
            Left  = Data[0,m]
            Right = Data[7,m]
            # plot results
            bw = 0.5
            ga = 0.3
        
            fig  = plt.figure(figsize=(5,8))
            ax1  = plt.axes([0.08,0.08,0.85,0.9])
        
            ProxyHandlesList = []   # For legend     
            # plot bars
            ax1.fill_between([0,0+bw], [0,0],[Left,Left],linestyle = '--', facecolor =MyColorCycle[11,:], linewidth = 0.0)
            ax1.fill_between([1,1+bw], [Data[1,m],Data[1,m]],[Left,Left],linestyle = '--', facecolor =MyColorCycle[4,:], linewidth = 0.0)
            ax1.fill_between([2,2+bw], [Data[2,m],Data[2,m]],[Data[1,m],Data[1,m]],linestyle = '--', facecolor =MyColorCycle[0,:], linewidth = 0.0)
            ax1.fill_between([3,3+bw], [Data[3,m],Data[3,m]],[Data[2,m],Data[2,m]],linestyle = '--', facecolor =MyColorCycle[18,:], linewidth = 0.0)
            ax1.fill_between([4,4+bw], [Data[4,m],Data[4,m]],[Data[3,m],Data[3,m]],linestyle = '--', facecolor =MyColorCycle[8,:], linewidth = 0.0)
            ax1.fill_between([5,5+bw], [Data[5,m],Data[5,m]],[Data[4,m],Data[4,m]],linestyle = '--', facecolor =MyColorCycle[16,:], linewidth = 0.0)
            ax1.fill_between([6,6+bw], [Data[6,m],Data[6,m]],[Data[5,m],Data[5,m]],linestyle = '--', facecolor =MyColorCycle[2,:], linewidth = 0.0)
            ax1.fill_between([7,7+bw], [Data[7,m],Data[7,m]],[Data[6,m],Data[6,m]],linestyle = '--', facecolor =MyColorCycle[6,:], linewidth = 0.0)
            ax1.fill_between([8,8+bw], [0,0],[Data[7,m],Data[7,m]],linestyle = '--', facecolor =MyColorCycle[15,:], linewidth = 0.0)
            
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[11,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[4,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[0,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[18,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[8,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[16,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[2,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[6,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[15,:])) # create proxy artist for legend
            
            # plot lines:
            plt.plot([0,8.5],[Left,Left],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([1,2.5],[Data[1,m],Data[1,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([2,3.5],[Data[2,m],Data[2,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([3,4.5],[Data[3,m],Data[3,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([4,5.5],[Data[4,m],Data[4,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([5,6.5],[Data[5,m],Data[5,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([6,7.5],[Data[6,m],Data[6,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([7,8.5],[Data[7,m],Data[7,m]],linestyle = '-', linewidth = 0.5, color = 'k')
    
            plt.arrow(8.25, Data[7,m],0, Data[0,m]-Data[7,m], lw = 0.8, ls = '-', shape = 'full',
                  length_includes_head = True, head_width =0.1, head_length =0.01*Left, ec = 'k', fc = 'k')
            plt.arrow(8.25,Data[0,m],0,Data[7,m]-Data[0,m], lw = 0.8, ls = '-', shape = 'full',
                  length_includes_head = True, head_width =0.1, head_length =0.01*Left, ec = 'k', fc = 'k')
    
            # plot text and labels
            plt.text(6.85, 0.94 *Left, ("%3.0f" % inc) + ' %',fontsize=18,fontweight='bold')          
            plt.text(4.3, 0.94  *Right, Scens[m],fontsize=18,fontweight='bold') 
            plt.title('RE strategies and GHG emissions, ' + Sector[0] + '.', fontsize = 18)
            plt.ylabel(Title[nn] + ', Mt.', fontsize = 18)
            plt.xticks([0.25,1.25,2.25,3.25,4.25,5.25,6.25,7.25,8.25])
            plt.yticks(fontsize =18)
            ax1.set_xticklabels([], rotation =90, fontsize = 21, fontweight = 'normal')
            plt_lgd  = plt.legend(handles = ProxyHandlesList,labels = LWE,shadow = False, prop={'size':12},ncol=1, loc = 'upper right' ,bbox_to_anchor=(1.91, 1)) 
            #plt.axis([-0.2, 7.7, 0.9*Right, 1.02*Left])
            plt.axis([-0.2, 8.7, 0, 1.02*Left])
        
            plt.show()
            fig_name = Title[nn] + Region + '_ ' + Sector[0] + '_' + Scens[m] + '.png'
            fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = 400, bbox_inches='tight')                 
    
    
    ### Area plot RE
            
    NS = 3
    NR = 2
    NE = 8
    Nt = 45
    Nm = 6
    
    AnnEmsV = np.zeros((Nt,NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    #AnnEmsB = np.zeros((Nt,NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario
    MatEmsV = np.zeros((Nt,NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    #MatEmsB = np.zeros((Nt,NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario
    MatProduction_Prim = np.zeros((Nt,Nm,NS,NR,NE))
    MatProduction_Sec  = np.zeros((Nt,Nm,NS,NR,NE))
    
    for r in range(0,NE): # RE scenario
        Path = os.path.join(RECC_Paths.results_path,FolderlistV[r],'SysVar_TotalGHGFootprint.xls')
        Resultfile   = xlrd.open_workbook(Path)
        Resultsheet  = Resultfile.sheet_by_name('TotalGHGFootprint')
        Resultsheet1 = Resultfile.sheet_by_name('Cover')
        UUID         = Resultsheet1.cell_value(3,2)
        Resultfile2  = xlrd.open_workbook(os.path.join(RECC_Paths.results_path,FolderlistV[r],'ODYM_RECC_ModelResults_' + UUID + '.xlsx'))
        Resultsheet2 = Resultfile2.sheet_by_name('Model_Results')
        # Find the index for the recycling credit and others:
        rci = 1
        while True:
            if Resultsheet2.cell_value(rci, 0) == 'GHG emissions, recycling credits':
                break # that gives us the right index to read the recycling credit from the result table.
            rci += 1
        mci = 1
        while True:
            if Resultsheet2.cell_value(mci, 0) == 'GHG emissions, material cycle industries and their energy supply _3di_9di':
                break # that gives us the right index to read the recycling credit from the result table.
            mci += 1
        mc1 = 1
        while True:
            if Resultsheet2.cell_value(mc1, 0) == 'Primary steel production':
                break # that gives us the right index from the result table.
            mc1 += 1
        mc2 = 1
        while True:
            if Resultsheet2.cell_value(mc2, 0) == 'Primary Al production':
                break # that gives us the right index from the result table.
            mc2 += 1
        mc3 = 1
        while True:
            if Resultsheet2.cell_value(mc3, 0) == 'Primary Cu production':
                break # that gives us the right index from the result table.
            mc3 += 1
        mc4 = 1
        while True:
            if Resultsheet2.cell_value(mc4, 0) == 'Cement production':
                break # that gives us the right index from the result table.
            mc4 += 1
        mc5 = 1
        while True:
            if Resultsheet2.cell_value(mc5, 0) == 'Primary plastics production':
                break # that gives us the right index from the result table.
            mc5 += 1
        mc6 = 1
        while True:
            if Resultsheet2.cell_value(mc6, 0) == 'Wood, from forests':
                break # that gives us the right index from the result table.
            mc6 += 1
        mc7 = 1
        while True:
            if Resultsheet2.cell_value(mc7, 0) == 'Secondary steel':
                break # that gives us the right index from the result table.
            mc7 += 1
        mc8 = 1
        while True:
            if Resultsheet2.cell_value(mc8, 0) == 'Secondary Al':
                break # that gives us the right index from the result table.
            mc8 += 1            
        mc9 = 1
        while True:
            if Resultsheet2.cell_value(mc9, 0) == 'Secondary copper':
                break # that gives us the right index from the result table.
            mc9 += 1   
        mc10 = 1
        while True:
            if Resultsheet2.cell_value(mc10, 0) == 'Secondary plastics':
                break # that gives us the right index from the result table.
            mc10 += 1   
        mc11 = 1
        while True:
            if Resultsheet2.cell_value(mc11, 0) == 'Recycled wood':
                break # that gives us the right index from the result table.
            mc11 += 1               
            
        ru1 = 1
        while True:
            if Resultsheet2.cell_value(ru1, 0) == 'ReUse of materials in products, construction grade steel':
                break # that gives us the right index from the result table.
            ru1 += 1 
        ru2 = 1
        while True:
            if Resultsheet2.cell_value(ru2, 0) == 'ReUse of materials in products, automotive steel':
                break # that gives us the right index from the result table.
            ru2 += 1             
        ru3 = 1
        while True:
            if Resultsheet2.cell_value(ru3, 0) == 'ReUse of materials in products, stainless steel':
                break # that gives us the right index from the result table.
            ru3 += 1 
        ru4 = 1
        while True:
            if Resultsheet2.cell_value(ru4, 0) == 'ReUse of materials in products, cast iron':
                break # that gives us the right index from the result table.
            ru4 += 1 
        ru5 = 1
        while True:
            if Resultsheet2.cell_value(ru5, 0) == 'ReUse of materials in products, wrought Al':
                break # that gives us the right index from the result table.
            ru5 += 1 
        ru6 = 1
        while True:
            if Resultsheet2.cell_value(ru6, 0) == 'ReUse of materials in products, cast Al':
                break # that gives us the right index from the result table.
            ru6 += 1
        ru7 = 1
        while True:
            if Resultsheet2.cell_value(ru7, 0) == 'ReUse of materials in products, copper electric grade':
                break # that gives us the right index from the result table.
            ru7 += 1
        ru8 = 1
        while True:
            if Resultsheet2.cell_value(ru8, 0) == 'ReUse of materials in products, plastics':
                break # that gives us the right index from the result table.
            ru8 += 1
        ru9 = 1
        while True:
            if Resultsheet2.cell_value(ru9, 0) == 'ReUse of materials in products, cement':
                break # that gives us the right index from the result table.
            ru9 += 1 
        ru10 = 1
        while True:
            if Resultsheet2.cell_value(ru10, 0) == 'ReUse of materials in products, wood and wood products':
                break # that gives us the right index from the result table.
            ru10 += 1            
         
            
        for s in range(0,NS): # SSP scenario
            for c in range(0,NR):
                for t in range(0,45): # time
                    AnnEmsV[t,s,c,r] = Resultsheet.cell_value(t +2, 1 + c + NR*s)
                    MatEmsV[t,s,c,r] = Resultsheet2.cell_value(mci+ 2*s +c,t+8)
        # Material results export
        for s in range(0,NS): # SSP scenario
            for c in range(0,NR):
                for t in range(0,35): # time until 2050 only!!! Cum. emissions until 2050.
                    MatCumEmsV[s,c,r] += Resultsheet2.cell_value(mci+ 2*s +c,t+8)
                for t in range(0,45): # time until 2060.
                    MatCumEmsV2060[s,c,r] += Resultsheet2.cell_value(mci+ 2*s +c,t+8)                    
                MatAnnEmsV2030[s,c,r]  = Resultsheet2.cell_value(mci+ 2*s +c,22)
                MatAnnEmsV2050[s,c,r]  = Resultsheet2.cell_value(mci+ 2*s +c,42)
            AvgDecadalMatEmsV[s,r,0]   = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for t in range(13,23)])/10
            AvgDecadalMatEmsV[s,r,1]   = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for t in range(23,33)])/10
            AvgDecadalMatEmsV[s,r,2]   = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for t in range(33,43)])/10
            AvgDecadalMatEmsV[s,r,3]   = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for t in range(43,53)])/10    
        # Material results export, including recycling credit
        for s in range(0,NS): # SSP scenario
            for c in range(0,NR):
                for t in range(0,35): # time until 2050 only!!! Cum. emissions until 2050.
                    MatCumEmsVC[s,c,r]+= Resultsheet2.cell_value(mci+ 2*s +c,t+8) + Resultsheet2.cell_value(rci+ 2*s +c,t+8)
                for t in range(0,45): # time until 2060.
                    MatCumEmsVC2060[s,c,r]+= Resultsheet2.cell_value(mci+ 2*s +c,t+8) + Resultsheet2.cell_value(rci+ 2*s +c,t+8)
                MatAnnEmsV2030C[s,c,r] = Resultsheet2.cell_value(mci+ 2*s +c,22)  + Resultsheet2.cell_value(rci+ 2*s +c,22)
                MatAnnEmsV2050C[s,c,r] = Resultsheet2.cell_value(mci+ 2*s +c,42)  + Resultsheet2.cell_value(rci+ 2*s +c,42)
            AvgDecadalMatEmsVC[s,r,0]  = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for t in range(13,23)])/10 + sum([Resultsheet2.cell_value(rci+ 2*s +1,t) for t in range(13,23)])/10
            AvgDecadalMatEmsVC[s,r,1]  = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for t in range(23,33)])/10 + sum([Resultsheet2.cell_value(rci+ 2*s +1,t) for t in range(23,33)])/10
            AvgDecadalMatEmsVC[s,r,2]  = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for t in range(33,43)])/10 + sum([Resultsheet2.cell_value(rci+ 2*s +1,t) for t in range(33,43)])/10
            AvgDecadalMatEmsVC[s,r,3]  = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for t in range(43,53)])/10 + sum([Resultsheet2.cell_value(rci+ 2*s +1,t) for t in range(43,53)])/10                       

        # Material results export, prim. and secondary prod.
        for s in range(0,NS): # SSP scenario
            for c in range(0,NR):
                for t in range(0,45): # time until 2060
                    MatProduction_Prim[t,0,s,c,r] = Resultsheet2.cell_value(mc1+ 2*s +c,t+8)
                    MatProduction_Prim[t,1,s,c,r] = Resultsheet2.cell_value(mc2+ 2*s +c,t+8)
                    MatProduction_Prim[t,2,s,c,r] = Resultsheet2.cell_value(mc3+ 2*s +c,t+8)
                    MatProduction_Prim[t,3,s,c,r] = Resultsheet2.cell_value(mc4+ 2*s +c,t+8)
                    MatProduction_Prim[t,4,s,c,r] = Resultsheet2.cell_value(mc5+ 2*s +c,t+8)
                    MatProduction_Prim[t,5,s,c,r] = Resultsheet2.cell_value(mc6+ 2*s +c,t+8)
                    
                    MatProduction_Sec[t,0,s,c,r]  = Resultsheet2.cell_value(mc7+ 2*s +c,t+8) + Resultsheet2.cell_value(ru1+ 2*s +c,t+8) + Resultsheet2.cell_value(ru2+ 2*s +c,t+8) + Resultsheet2.cell_value(ru3+ 2*s +c,t+8) + Resultsheet2.cell_value(ru4+ 2*s +c,t+8)
                    MatProduction_Sec[t,1,s,c,r]  = Resultsheet2.cell_value(mc8+ 2*s +c,t+8) + Resultsheet2.cell_value(ru5+ 2*s +c,t+8) + Resultsheet2.cell_value(ru6+ 2*s +c,t+8)
                    MatProduction_Sec[t,2,s,c,r]  = Resultsheet2.cell_value(mc9+ 2*s +c,t+8) + Resultsheet2.cell_value(ru7+ 2*s +c,t+8)
                    MatProduction_Sec[t,3,s,c,r]  = Resultsheet2.cell_value(ru9+ 2*s +c,t+8)
                    MatProduction_Sec[t,4,s,c,r]  = Resultsheet2.cell_value(mc10+ 2*s +c,t+8) + Resultsheet2.cell_value(ru8+ 2*s +c,t+8)
                    MatProduction_Sec[t,5,s,c,r]  = Resultsheet2.cell_value(mc11+ 2*s +c,t+8) + Resultsheet2.cell_value(ru10+ 2*s +c,t+8)

    MatSummaryV[0:3,:] = MatAnnEmsV2030[:,1,:].copy() # RCP is fixed: RCP2.6
    MatSummaryV[3:6,:] = MatAnnEmsV2050[:,1,:].copy() # RCP is fixed: RCP2.6
    MatSummaryV[6:9,:] = MatCumEmsV[:,1,:].copy()     # RCP is fixed: RCP2.6
    MatSummaryV[9::,:] = MatCumEmsV2060[:,1,:].copy() # RCP is fixed: RCP2.6
    
    MatSummaryVC[0:3,:]= MatAnnEmsV2030C[:,1,:].copy() # RCP is fixed: RCP2.6
    MatSummaryVC[3:6,:]= MatAnnEmsV2050C[:,1,:].copy() # RCP is fixed: RCP2.6
    MatSummaryVC[6:9,:]= MatCumEmsVC[:,1,:].copy()     # RCP is fixed: RCP2.6
    MatSummaryVC[9::,:]= MatCumEmsVC2060[:,1,:].copy() # RCP is fixed: RCP2.6
    
    # Area plot, stacked, GHG emissions, system
    MyColorCycle = pylab.cm.Set1(np.arange(0,1,0.1)) # select colors from the 'Paired' color map.            
    grey0_9      = np.array([0.9,0.9,0.9,1])
    
    Title      = ['GHG_System_RES_stack','GHG_material_cycles_RES_stack']
    Sector     = ['pav_reb_nrb']
    Scens      = ['LED','SSP1','SSP2']
    LWE_area   = ['higher yields', 're-use & LTE','material subst.','down-sizing','car-sharing','ride-sharing','More intense bld. use']     
    
    for nn in range(0,len(Title)):
        #mS = 1
        #mR = 1
        mRCP = 1 # select RCP2.6, which has full implementation of RE strategies by 2050.
        for mS in range(0,NS): # SSP
            for mR in range(0,1): # Vehs
                
                if nn == 0 and mR == 0:
                    Data = AnnEmsV[:,mS,mRCP,:]
                
                if nn == 1 and mR == 0:
                    Data = MatEmsV[:,mS,mRCP,:]                
                
                fig  = plt.figure(figsize=(8,5))
                ax1  = plt.axes([0.08,0.08,0.85,0.9])
                
                ProxyHandlesList = []   # For legend     
                
                # plot area
                ax1.fill_between(np.arange(2016,2061),np.zeros((Nt)), Data[:,-1], linestyle = '-', facecolor = grey0_9, linewidth = 1.0, alpha=0.5)
                ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=grey0_9)) # create proxy artist for legend
                for m in range(7,0,-1):
                    ax1.fill_between(np.arange(2016,2061),Data[:,m], Data[:,m-1], linestyle = '-', facecolor = MyColorCycle[m,:], linewidth = 1.0, alpha=0.5)
                    ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:], alpha=0.75)) # create proxy artist for legend
                    ax1.plot(np.arange(2016,2061),Data[:,m],linestyle = '--', color = MyColorCycle[m,:], linewidth = 1.1,)                
                ax1.plot(np.arange(2016,2061),Data[:,0],linestyle = '--', color = 'k', linewidth = 1.1,)               
                #plt.text(Data[m,:].min()*0.55, 7.8, 'Baseline: ' + ("%3.0f" % Base[m]) + ' Mt/yr.',fontsize=14,fontweight='bold')
                plt.text(2027,Data[m,:].max()*1.02, 'Colors may deviate from legend colors due to overlap of RES wedges.',fontsize=8.5,fontweight='bold')
                
                plt.title(Title[nn] + ' \n' + Region + ', ' + Sector[mR] + ', ' + Scens[mS] + '.', fontsize = 18)
                plt.ylabel('Mt of CO2-eq.', fontsize = 18)
                plt.xlabel('Year', fontsize = 18)
                plt.xticks(fontsize=18)
                plt.yticks(fontsize=18)
                if mR == 0: # vehicles, legend lower left
                    plt_lgd  = plt.legend(handles = reversed(ProxyHandlesList),labels = LWE_area, shadow = False, prop={'size':12},ncol=1, loc = 'lower left')# ,bbox_to_anchor=(1.91, 1)) 
                ax1.set_xlim([2015, 2061])
                
                plt.show()
                fig_name = Title[nn] + '_' + Region + '_' + Sector[mR] + '_' + Scens[mS] + '.png'
                fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = 400, bbox_inches='tight')            
                
    ##### Overview plot metal production
    MyColorCycle = pylab.cm.tab20(np.arange(0,1,0.05)) # select colors from the 'tab20' color map.            
    grey0_9      = np.array([0.9,0.9,0.9,1])
    
    Title      = ['Materials']
    Sector     = ['pav_reb_nrb']
    Scens      = ['LED','SSP1','SSP2']
    
    bw = 0.7
    #mS = 0
    #mR = 0
    mRCP = 1 # select RCP2.6, which has full implementation of RE strategies by 2050.
    for mS in range(0,NS): # SSP
        for mR in range(0,1): # pav-reb-nrb
                          
            fig, ((ax1, ax2, ax3), (ax4, ax5, ax6)) = plt.subplots(2, 3, sharex=True, gridspec_kw={'hspace': 0.3, 'wspace': 0.35})
            
            ax1.fill_between([1,1+bw], [0,0],[MatProduction_Prim[4,0,mS,1,0],MatProduction_Prim[4,0,mS,1,0]],linestyle = '--', facecolor =MyColorCycle[0,:], linewidth = 0.0)
            ax1.fill_between([1.8,1.8+bw], [0,0],[MatProduction_Prim[24:34,0,mS,1,0].sum()/10,MatProduction_Prim[24:34,0,mS,1,0].sum()/10],linestyle = '--', facecolor =MyColorCycle[0,:], linewidth = 0.0)
            ax1.fill_between([2.6,2.6+bw], [0,0],[MatProduction_Prim[24:34,0,mS,1,-1].sum()/10,MatProduction_Prim[24:34,0,mS,1,-1].sum()/10],linestyle = '--', facecolor =MyColorCycle[0,:], linewidth = 0.0)
            ax1.fill_between([4,4+bw], [0,0],[MatProduction_Sec[4,0,mS,1,0],MatProduction_Sec[4,0,mS,1,0]],linestyle = '--', facecolor =MyColorCycle[1,:], linewidth = 0.0)
            ax1.fill_between([4.8,4.8+bw], [0,0],[MatProduction_Sec[24:34,0,mS,1,0].sum()/10,MatProduction_Sec[24:34,0,mS,1,0].sum()/10],linestyle = '--', facecolor =MyColorCycle[1,:], linewidth = 0.0)
            ax1.fill_between([5.6,5.6+bw], [0,0],[MatProduction_Sec[24:34,0,mS,1,-1].sum()/10,MatProduction_Sec[24:34,0,mS,1,-1].sum()/10],linestyle = '--', facecolor =MyColorCycle[1,:], linewidth = 0.0)
            ax1.set_title('Steel')
            ax2.fill_between([1,1+bw], [0,0],[MatProduction_Prim[4,1,mS,1,0],MatProduction_Prim[4,1,mS,1,0]],linestyle = '--', facecolor =MyColorCycle[2,:], linewidth = 0.0)
            ax2.fill_between([1.8,1.8+bw], [0,0],[MatProduction_Prim[24:34,1,mS,1,0].sum()/10,MatProduction_Prim[24:34,1,mS,1,0].sum()/10],linestyle = '--', facecolor =MyColorCycle[2,:], linewidth = 0.0)
            ax2.fill_between([2.6,2.6+bw], [0,0],[MatProduction_Prim[24:34,1,mS,1,-1].sum()/10,MatProduction_Prim[24:34,1,mS,1,-1].sum()/10],linestyle = '--', facecolor =MyColorCycle[2,:], linewidth = 0.0)
            ax2.fill_between([4,4+bw], [0,0],[MatProduction_Sec[4,1,mS,1,0],MatProduction_Sec[4,1,mS,1,0]],linestyle = '--', facecolor =MyColorCycle[3,:], linewidth = 0.0)
            ax2.fill_between([4.8,4.8+bw], [0,0],[MatProduction_Sec[24:34,1,mS,1,0].sum()/10,MatProduction_Sec[24:34,1,mS,1,0].sum()/10],linestyle = '--', facecolor =MyColorCycle[3,:], linewidth = 0.0)
            ax2.fill_between([5.6,5.6+bw], [0,0],[MatProduction_Sec[24:34,1,mS,1,-1].sum()/10,MatProduction_Sec[24:34,1,mS,1,-1].sum()/10],linestyle = '--', facecolor =MyColorCycle[3,:], linewidth = 0.0)
            ax2.set_title('Aluminium')
            ax3.fill_between([1,1+bw], [0,0],[MatProduction_Prim[4,2,mS,1,0],MatProduction_Prim[4,2,mS,1,0]],linestyle = '--', facecolor =MyColorCycle[4,:], linewidth = 0.0)
            ax3.fill_between([1.8,1.8+bw], [0,0],[MatProduction_Prim[24:34,2,mS,1,0].sum()/10,MatProduction_Prim[24:34,2,mS,1,0].sum()/10],linestyle = '--', facecolor =MyColorCycle[4,:], linewidth = 0.0)
            ax3.fill_between([2.6,2.6+bw], [0,0],[MatProduction_Prim[24:34,2,mS,1,-1].sum()/10,MatProduction_Prim[24:34,2,mS,1,-1].sum()/10],linestyle = '--', facecolor =MyColorCycle[4,:], linewidth = 0.0)
            ax3.fill_between([4,4+bw], [0,0],[MatProduction_Sec[4,2,mS,1,0],MatProduction_Sec[4,2,mS,1,0]],linestyle = '--', facecolor =MyColorCycle[5,:], linewidth = 0.0)
            ax3.fill_between([4.8,4.8+bw], [0,0],[MatProduction_Sec[24:34,2,mS,1,0].sum()/10,MatProduction_Sec[24:34,2,mS,1,0].sum()/10],linestyle = '--', facecolor =MyColorCycle[5,:], linewidth = 0.0)
            ax3.fill_between([5.6,5.6+bw], [0,0],[MatProduction_Sec[24:34,2,mS,1,-1].sum()/10,MatProduction_Sec[24:34,2,mS,1,-1].sum()/10],linestyle = '--', facecolor =MyColorCycle[5,:], linewidth = 0.0)
            ax3.set_title('Copper')
            ax4.fill_between([1,1+bw], [0,0],[MatProduction_Prim[4,3,mS,1,0],MatProduction_Prim[4,3,mS,1,0]],linestyle = '--', facecolor =MyColorCycle[6,:], linewidth = 0.0)
            ax4.fill_between([1.8,1.8+bw], [0,0],[MatProduction_Prim[24:34,3,mS,1,0].sum()/10,MatProduction_Prim[24:34,3,mS,1,0].sum()/10],linestyle = '--', facecolor =MyColorCycle[6,:], linewidth = 0.0)
            ax4.fill_between([2.6,2.6+bw], [0,0],[MatProduction_Prim[24:34,3,mS,1,-1].sum()/10,MatProduction_Prim[24:34,3,mS,1,-1].sum()/10],linestyle = '--', facecolor =MyColorCycle[6,:], linewidth = 0.0)
            ax4.fill_between([4,4+bw], [0,0],[MatProduction_Sec[4,3,mS,1,0],MatProduction_Sec[4,3,mS,1,0]],linestyle = '--', facecolor =MyColorCycle[7,:], linewidth = 0.0)
            ax4.fill_between([4.8,4.8+bw], [0,0],[MatProduction_Sec[24:34,3,mS,1,0].sum()/10,MatProduction_Sec[24:34,3,mS,1,0].sum()/10],linestyle = '--', facecolor =MyColorCycle[7,:], linewidth = 0.0)
            ax4.fill_between([5.6,5.6+bw], [0,0],[MatProduction_Sec[24:34,3,mS,1,-1].sum()/10,MatProduction_Sec[24:34,3,mS,1,-1].sum()/10],linestyle = '--', facecolor =MyColorCycle[7,:], linewidth = 0.0)
            ax4.set_title('Cement')
            ax5.fill_between([1,1+bw], [0,0],[MatProduction_Prim[4,4,mS,1,0],MatProduction_Prim[4,4,mS,1,0]],linestyle = '--', facecolor =MyColorCycle[8,:], linewidth = 0.0)
            ax5.fill_between([1.8,1.8+bw], [0,0],[MatProduction_Prim[24:34,4,mS,1,0].sum()/10,MatProduction_Prim[24:34,4,mS,1,0].sum()/10],linestyle = '--', facecolor =MyColorCycle[8,:], linewidth = 0.0)
            ax5.fill_between([2.6,2.6+bw], [0,0],[MatProduction_Prim[24:34,4,mS,1,-1].sum()/10,MatProduction_Prim[24:34,4,mS,1,-1].sum()/10],linestyle = '--', facecolor =MyColorCycle[8,:], linewidth = 0.0)
            ax5.fill_between([4,4+bw], [0,0],[MatProduction_Sec[4,4,mS,1,0],MatProduction_Sec[4,4,mS,1,0]],linestyle = '--', facecolor =MyColorCycle[9,:], linewidth = 0.0)
            ax5.fill_between([4.8,4.8+bw], [0,0],[MatProduction_Sec[24:34,4,mS,1,0].sum()/10,MatProduction_Sec[24:34,4,mS,1,0].sum()/10],linestyle = '--', facecolor =MyColorCycle[9,:], linewidth = 0.0)
            ax5.fill_between([5.6,5.6+bw], [0,0],[MatProduction_Sec[24:34,4,mS,1,-1].sum()/10,MatProduction_Sec[24:34,4,mS,1,-1].sum()/10],linestyle = '--', facecolor =MyColorCycle[9,:], linewidth = 0.0)
            ax5.set_title('Plastics')
            ax6.fill_between([1,1+bw], [0,0],[MatProduction_Prim[4,5,mS,1,0],MatProduction_Prim[4,5,mS,1,0]],linestyle = '--', facecolor =MyColorCycle[10,:], linewidth = 0.0)
            ax6.fill_between([1.8,1.8+bw], [0,0],[MatProduction_Prim[24:34,5,mS,1,0].sum()/10,MatProduction_Prim[24:34,5,mS,1,0].sum()/10],linestyle = '--', facecolor =MyColorCycle[10,:], linewidth = 0.0)
            ax6.fill_between([2.6,2.6+bw], [0,0],[MatProduction_Prim[24:34,5,mS,1,-1].sum()/10,MatProduction_Prim[24:34,5,mS,1,-1].sum()/10],linestyle = '--', facecolor =MyColorCycle[10,:], linewidth = 0.0)
            ax6.fill_between([4,4+bw], [0,0],[MatProduction_Sec[4,5,mS,1,0],MatProduction_Sec[4,5,mS,1,0]],linestyle = '--', facecolor =MyColorCycle[11,:], linewidth = 0.0)
            ax6.fill_between([4.8,4.8+bw], [0,0],[MatProduction_Sec[24:34,5,mS,1,0].sum()/10,MatProduction_Sec[24:34,5,mS,1,0].sum()/10],linestyle = '--', facecolor =MyColorCycle[11,:], linewidth = 0.0)
            ax6.fill_between([5.6,5.6+bw], [0,0],[MatProduction_Sec[24:34,5,mS,1,-1].sum()/10,MatProduction_Sec[24:34,5,mS,1,-1].sum()/10],linestyle = '--', facecolor =MyColorCycle[11,:], linewidth = 0.0)
            ax6.set_title('Wood')

            plt.sca(ax4)
            plt.xticks([1.4,2.2,3.0,4.4,5.2,6.0], ['2020','2040-50, no ME','2040-50, ME','2020','2040-50, no ME','2040-50, ME'], rotation =90, fontsize = 10, fontweight = 'normal')
            plt.sca(ax5)
            plt.xticks([1.4,2.2,3.0,4.4,5.2,6.0], ['2020','2040-50, no ME','2040-50, ME','2020','2040-50, no ME','2040-50, ME'], rotation =90, fontsize = 10, fontweight = 'normal')
            plt.sca(ax6)
            plt.xticks([1.4,2.2,3.0,4.4,5.2,6.0], ['2020','2040-50, no ME','2040-50, ME','2020','2040-50, no ME','2040-50, ME'], rotation =90, fontsize = 10, fontweight = 'normal')

            plt.sca(ax1)
            plt.ylabel('Mt/yr', fontsize = 12)
            plt.sca(ax4)
            plt.ylabel('Mt/yr', fontsize = 12)
            
#            plt.text(2027,Data[m,:].max()*1.02, 'Colors may deviate from legend colors due to overlap of RES wedges.',fontsize=8.5,fontweight='bold')
#            #            plt.title(Title[nn] + ' \n' + Region + ', ' + Sector[mR] + ', ' + Scens[mS] + '.', fontsize = 18)
#            plt.ylabel('Mt of CO2-eq.', fontsize = 18)
#            plt.xlabel('Year', fontsize = 18)
#            plt.xticks(fontsize=18)
#            plt.yticks(fontsize=18)
#            if mR == 0: # vehicles, legend lower left
#                plt_lgd  = plt.legend(handles = reversed(ProxyHandlesList),labels = LWE_area, shadow = False, prop={'size':12},ncol=1, loc = 'lower left')# ,bbox_to_anchor=(1.91, 1)) 
#            ax1.set_xlim([2015, 2061])
            
            plt.show()
            fig_name = Title[0] + '_' + Region + '_' + Sector[mR] + '_' + Scens[mS] + '.png'
            fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = 400, bbox_inches='tight')                   
             
    ##### line Plot overview of primary steel and steel recycling
    
    # Select scenario list: same as for bar chart above
    # E.g. for the USA, run code lines 41 to 59.
    
    MyColorCycle = pylab.cm.Paired(np.arange(0,1,0.2))
    #linewidth = [1.2,2.4,1.2,1.2,1.2]
    linewidth  = [1.2,2,1.2]
    linewidth2 = [1.2,2,1.2]
    
    ColorOrder         = [1,0,3]
            
    NS = 3
    NR = 2
    NE = 8
    Nt = 45
    
    # Primary steel
    AnnEmsV_PrimarySteel   = np.zeros((Nt,NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    AnnEmsV_SecondarySteel = np.zeros((Nt,NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    
    for r in range(0,NE): # RE scenario
        Path         = os.path.join(RECC_Paths.results_path,FolderlistV[r],'SysVar_TotalGHGFootprint.xls')
        Resultfile1  = xlrd.open_workbook(Path)
        Resultsheet1 = Resultfile1.sheet_by_name('Cover')
        UUID         = Resultsheet1.cell_value(3,2)
        Resultfile2  = xlrd.open_workbook(os.path.join(RECC_Paths.results_path,FolderlistV[r],'ODYM_RECC_ModelResults_' + UUID + '.xlsx'))
        Resultsheet2 = Resultfile2.sheet_by_name('Model_Results')
        for s in range(0,NS): # SSP scenario
            for c in range(0,NR):
                for t in range(0,45): # timeAnnEmsV_SecondarySteel[t,s,c,r] = Resultsheet2.cell_value(151+ 2*s +c,t+8)
                    AnnEmsV_PrimarySteel[t,s,c,r] = Resultsheet2.cell_value(19+ 2*s +c,t+8)
                    AnnEmsV_SecondarySteel[t,s,c,r] = Resultsheet2.cell_value(163+ 2*s +c,t+8)
                    
    Title      = ['primary_steel','secondary_steel']            
    Sector     = ['pav_reb_nrb']
    ScensL     = ['SSP2, no REFs','SSP2, full REF spectrum','SSP1, no REFs','SSP1, full REF spectrum','LED, no REFs','LED, full REF spectrum']
    
    #mS = 1
    #mR = 1
    for nn in range(0,2):
        mRCP = 1 # select RCP2.6, which has full implementation of RE strategies by 2050.
        for mR in range(0,1): # Veh/Buildings
            
            if nn == 0:
                Data = AnnEmsV_PrimarySteel[:,:,mRCP,:]
            if nn == 1:
                Data = AnnEmsV_SecondarySteel[:,:,mRCP,:]
        
        
            fig  = plt.figure(figsize=(8,5))
            ax1  = plt.axes([0.08,0.08,0.85,0.9])
            
            ProxyHandlesList = []   # For legend     
            
            for mS in range(NS-1,-1,-1):
                ax1.plot(np.arange(2016,2061), Data[:,mS,0],  linewidth = linewidth[mS],  linestyle = '-',  color = MyColorCycle[ColorOrder[mS],:])
                #ProxyHandlesList.append(plt.line((0, 0), 1, 1, fc=MyColorCycle[m,:]))
                ax1.plot(np.arange(2016,2061), Data[:,mS,-1], linewidth = linewidth2[mS], linestyle = '--', color = MyColorCycle[ColorOrder[mS],:])
                #ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))     
            plt_lgd  = plt.legend(ScensL,shadow = False, prop={'size':12}, loc = 'upper left',bbox_to_anchor=(1.05, 1))    
            plt.ylabel('Mt/yr.', fontsize = 18) 
            plt.xlabel('year', fontsize = 18)         
            plt.title(Title[nn] + ', by socio-economic scenario, \n' + Region + ', ' + Sector[mR] + '.', fontsize = 18)
            plt.xticks(fontsize=18)
            plt.yticks(fontsize=18)
            ax1.set_xlim([2015, 2061])
            plt.gca().set_ylim(bottom=0)
            
            plt.show()
            fig_name = Title[nn] + '_' + Region + '_ ' + Sector[mR] + '.png'
            fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = 400, bbox_inches='tight')             
        
    
    return ASummaryV, AvgDecadalEmsV, MatSummaryV, AvgDecadalMatEmsV, MatSummaryVC, AvgDecadalMatEmsVC, MatProduction_Prim, MatProduction_Sec

# code for script to be run as standalone function
if __name__ == "__main__":
    main()

#
#
#
#
#