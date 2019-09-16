# -*- coding: utf-8 -*-
"""
Created on Wed Oct 17 10:37:00 2018

@author: spauliuk
"""
def main(RegionalScope,ResBldgsList):
    
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
    # 6) + EoL + FSD + FYI + ReU +LTE + MSu + LWE + MIU = ALL 
    
    Region      = RegionalScope
    FolderlistB = ResBldgsList
    
    # Waterfall plots.
    
    NS = 3 # no of SSP scenarios
    NR = 2 # no of RCP scenarios
    NE = 6 # no of Res. eff. scenarios
    
    CumEmsV        = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    AnnEmsV2030    = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    AnnEmsV2050    = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    ASummaryV      = np.zeros((9,NE)) # For direct copy-paste to Excel
    AvgDecadalEmsV = np.zeros((NS,NE,4)) # SSP-Scenario x RES scenario, RCP fixed to RCP2.6
    # for materials:
    MatCumEmsV        = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    MatAnnEmsV2030    = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    MatAnnEmsV2050    = np.zeros((NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario    
    MatSummaryV       = np.zeros((9,NE)) # For direct copy-paste to Excel
    AvgDecadalMatEmsV = np.zeros((NS,NE,4)) # SSP-Scenario x RES scenario, RCP is fixed: RCP2.6
    
    for r in range(0,NE): # RE scenario
        Path = os.path.join(RECC_Paths.results_path,FolderlistB[r],'SysVar_TotalGHGFootprint.xls')
        Resultfile = xlrd.open_workbook(Path)
        Resultsheet = Resultfile.sheet_by_name('TotalGHGFootprint')
        for s in range(0,NS): # SSP scenario
            for c in range(0,NR):
                for t in range(0,35): # time until 2050 only!!! Cum. emissions until 2050.
                    CumEmsV[s,c,r] += Resultsheet.cell_value(t +2, 1 + c + NR*s)
                AnnEmsV2030[s,c,r]  = Resultsheet.cell_value(16  , 1 + c + NR*s)
                AnnEmsV2050[s,c,r]  = Resultsheet.cell_value(36  , 1 + c + NR*s)
            AvgDecadalEmsV[s,r,0]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(7,17)])/10
            AvgDecadalEmsV[s,r,1]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(17,27)])/10
            AvgDecadalEmsV[s,r,2]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(27,37)])/10
            AvgDecadalEmsV[s,r,3]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(37,47)])/10                    
                
    ASummaryV[0:3,:] = AnnEmsV2030[:,1,:].copy()
    ASummaryV[3:6,:] = AnnEmsV2050[:,1,:].copy()
    ASummaryV[6::,:] = CumEmsV[:,1,:].copy()
                        
    # Waterfall plot            
    MyColorCycle = pylab.cm.Set1(np.arange(0,1,0.14)) # select 12 colors from the 'Paired' color map.            
    
    Sector = ['Residential_buildings']
    Title  = ['Cum_GHG_2016_2050','Cum_GHG_2040_2050','Annual_GHG_2050']
    Scens  = ['LED','SSP1','SSP2']
    LWE    = ['No RE','higher yields', 're-use/longer use','material subst.','light design','more intense use','All RE stratgs.']
    
    for nn in range(0,3):
        for m in range(0,NS): # SSP
            if nn == 0:
                Data = np.einsum('SE->ES',CumEmsV[:,1,:])
            if nn == 1:
                Data = np.einsum('SE->ES',10*AvgDecadalEmsV[:,:,2])
            if nn == 2:
                Data = np.einsum('SE->ES',AnnEmsV2050[:,1,:])
                
            inc = -100 * (Data[0,m] - Data[5,m])/Data[0,m]
        
            Left  = Data[0,m]
            Right = Data[5,m]
            # plot results
            bw = 0.5
            ga = 0.3
        
            fig  = plt.figure(figsize=(5,8))
            ax1  = plt.axes([0.08,0.08,0.85,0.9])
        
            ProxyHandlesList = []   # For legend     
            # plot bars
            ax1.fill_between([0,0+bw], [0,0],[Left,Left],linestyle = '--', facecolor =MyColorCycle[0,:], linewidth = 0.0)
            ax1.fill_between([1,1+bw], [Data[1,m],Data[1,m]],[Left,Left],linestyle = '--', facecolor =MyColorCycle[1,:], linewidth = 0.0)
            ax1.fill_between([2,2+bw], [Data[2,m],Data[2,m]],[Data[1,m],Data[1,m]],linestyle = '--', facecolor =MyColorCycle[2,:], linewidth = 0.0)
            ax1.fill_between([3,3+bw], [Data[3,m],Data[3,m]],[Data[2,m],Data[2,m]],linestyle = '--', facecolor =MyColorCycle[3,:], linewidth = 0.0)
            ax1.fill_between([4,4+bw], [Data[4,m],Data[4,m]],[Data[3,m],Data[3,m]],linestyle = '--', facecolor =MyColorCycle[4,:], linewidth = 0.0)
            ax1.fill_between([5,5+bw], [Data[5,m],Data[5,m]],[Data[4,m],Data[4,m]],linestyle = '--', facecolor =MyColorCycle[5,:], linewidth = 0.0)
            ax1.fill_between([6,6+bw], [0,0],[Data[5,m],Data[5,m]],linestyle = '--', facecolor =MyColorCycle[6,:], linewidth = 0.0)
            
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[0,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[1,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[2,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[3,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[4,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[5,:])) # create proxy artist for legend
            ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[6,:])) # create proxy artist for legend
            
            # plot lines:
            plt.plot([0,7.5],[Left,Left],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([1,2.5],[Data[1,m],Data[1,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([2,3.5],[Data[2,m],Data[2,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([3,4.5],[Data[3,m],Data[3,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([4,5.5],[Data[4,m],Data[4,m]],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([5,6.5],[Data[5,m],Data[5,m]],linestyle = '-', linewidth = 0.5, color = 'k')
    
            plt.arrow(6.25, Data[5,m],0, Data[0,m]-Data[5,m], lw = 0.8, ls = '-', shape = 'full',
                  length_includes_head = True, head_width =0.1, head_length =0.01*Left, ec = 'k', fc = 'k')
            plt.arrow(6.25,Data[0,m],0,Data[5,m]-Data[0,m], lw = 0.8, ls = '-', shape = 'full',
                  length_includes_head = True, head_width =0.1, head_length =0.01*Left, ec = 'k', fc = 'k')
    
            # plot text and labels
            plt.text(5.00, 0.94 *Left, ("%3.0f" % inc) + ' %',fontsize=18,fontweight='bold')          
            plt.text(2.8, 0.94  *Right, Scens[m],fontsize=18,fontweight='bold') 
            plt.title('RE strategies and GHG emissions, ' + Sector[0] + '.', fontsize = 18)
            plt.ylabel(Title[nn] + ', Mt.', fontsize = 18)
            plt.xticks([0.25,1.25,2.25,3.25,4.25,5.25,6.25])
            plt.yticks(fontsize =18)
            ax1.set_xticklabels([], rotation =90, fontsize = 21, fontweight = 'normal')
            plt_lgd  = plt.legend(handles = ProxyHandlesList,labels = LWE,shadow = False, prop={'size':12},ncol=1, loc = 'upper right' ,bbox_to_anchor=(1.91, 1)) 
            #plt.axis([-0.2, 7.7, 0.9*Right, 1.02*Left])
            plt.axis([-0.2, 6.7, 0, 1.02*Left])
        
            plt.show()
            fig_name = Title[nn] + Region + '_ ' + Sector[0] + '_' + Scens[m] + '.png'
            fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = 400, bbox_inches='tight')             
                
    
    ### Area plot RE
            
    NS = 3
    NR = 2
    NE = 6
    Nt = 45
    
    AnnEmsV = np.zeros((Nt,NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    #AnnEmsB = np.zeros((Nt,NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario
    MatEmsV = np.zeros((Nt,NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    #MatEmsB = np.zeros((Nt,NS,NC,NR)) # SSP-Scenario x RCP scenario x RES scenario
    
    
    for r in range(0,NE): # RE scenario
        Path = os.path.join(RECC_Paths.results_path,FolderlistB[r],'SysVar_TotalGHGFootprint.xls')
        Resultfile   = xlrd.open_workbook(Path)
        Resultsheet  = Resultfile.sheet_by_name('TotalGHGFootprint')
        Resultsheet1 = Resultfile.sheet_by_name('Cover')
        UUID         = Resultsheet1.cell_value(3,2)
        Resultfile2  = xlrd.open_workbook(os.path.join(RECC_Paths.results_path,FolderlistB[r],'ODYM_RECC_ModelResults_' + UUID + '.xls'))
        Resultsheet2 = Resultfile2.sheet_by_name('Model_Results')
        for s in range(0,NS): # SSP scenario
            for c in range(0,NR):
                for t in range(0,45): # time
                    AnnEmsV[t,s,c,r] = Resultsheet.cell_value(t +2, 1 + c + NR*s)
                    MatEmsV[t,s,c,r] = Resultsheet2.cell_value(229+ 2*s +c,t+8)
        # Material results export
        for s in range(0,NS): # SSP scenario
            for c in range(0,NR):
                for t in range(0,35): # time until 2050 only!!! Cum. emissions until 2050.
                    MatCumEmsV[s,c,r] += Resultsheet2.cell_value(229+ 2*s +c,t+8)
                MatAnnEmsV2030[s,c,r]  = Resultsheet2.cell_value(229+ 2*s +c,22)
                MatAnnEmsV2050[s,c,r]  = Resultsheet2.cell_value(229+ 2*s +c,42)
            AvgDecadalMatEmsV[s,r,0]   = sum([Resultsheet2.cell_value(229+ 2*s +1,t) for i in range(12,22)])/10
            AvgDecadalMatEmsV[s,r,1]   = sum([Resultsheet2.cell_value(229+ 2*s +1,t) for i in range(22,32)])/10
            AvgDecadalMatEmsV[s,r,2]   = sum([Resultsheet2.cell_value(229+ 2*s +1,t) for i in range(32,42)])/10
            AvgDecadalMatEmsV[s,r,3]   = sum([Resultsheet2.cell_value(229+ 2*s +1,t) for i in range(42,52)])/10    

    MatSummaryV[0:3,:] = MatAnnEmsV2030[:,1,:].copy() # RCP is fixed: RCP2.6
    MatSummaryV[3:6,:] = MatAnnEmsV2050[:,1,:].copy() # RCP is fixed: RCP2.6
    MatSummaryV[6::,:] = MatCumEmsV[:,1,:].copy()     # RCP is fixed: RCP2.6                    
    
    # Area plot, stacked, GHG emissions, system
    MyColorCycle = pylab.cm.Set1(np.arange(0,1,0.1)) # select colors from the 'Paired' color map.            
    grey0_9      = np.array([0.9,0.9,0.9,1])
    
    Title      = ['GHG_System_RES_stack','GHG_material_cycles_RES_stack']
    Sector     = ['residential_buildings']
    Scens      = ['LED','SSP1','SSP2']
    LWE_area   = ['higher yields', 're-use & LTE','material subst.','light design','more intense use']     
    
    for nn in range(0,len(Title)):
        #mS = 1
        #mR = 1
        mRCP = 1 # select RCP2.6, which has full implementation of RE strategies by 2050.
        for mS in range(0,NS): # SSP
            for mR in range(0,1): # Blds
                
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
                for m in range(5,0,-1):
                    ax1.fill_between(np.arange(2016,2061),Data[:,m], Data[:,m-1], linestyle = '-', facecolor = MyColorCycle[m,:], linewidth = 1.0, alpha=0.5)
                    ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:], alpha=0.75)) # create proxy artist for legend
                    ax1.plot(np.arange(2016,2061),Data[:,m],linestyle = '--', color = MyColorCycle[m,:], linewidth = 0.8,)                
                ax1.plot(np.arange(2016,2061),Data[:,0],linestyle = '--', color = 'k', linewidth = 0.8,)               
                #plt.text(Data[m,:].min()*0.55, 7.8, 'Baseline: ' + ("%3.0f" % Base[m]) + ' Mt/yr.',fontsize=14,fontweight='bold')
                plt.text(2027,Data[m,:].max()*1.02, 'Colors may deviate from legend colors due to overlap of RES wedges.',fontsize=8.5,fontweight='bold')
                
                plt.title(Title[nn] + ' \n' + Region + ', ' + Sector[mR] + ', ' + Scens[mS] + '.', fontsize = 18)
                plt.ylabel('Mt of CO2-eq.', fontsize = 18)
                plt.xlabel('Year', fontsize = 18)
                plt.xticks(fontsize=18)
                plt.yticks(fontsize=18)
                if mR == 0: # buildings, upper right
                    plt_lgd  = plt.legend(handles = reversed(ProxyHandlesList),labels = LWE_area, shadow = False, prop={'size':12},ncol=1, loc = 'upper right')# ,bbox_to_anchor=(1.91, 1)) 
                ax1.set_xlim([2015, 2061])
                
                plt.show()
                fig_name = Title[nn] + '_' + Region + '_' + Sector[mR] + '_' + Scens[mS] + '.png'
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
    NE = 6
    Nt = 45
    
    # Primary steel
    AnnEmsV_PrimarySteel   = np.zeros((Nt,NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    AnnEmsV_SecondarySteel = np.zeros((Nt,NS,NR,NE)) # SSP-Scenario x RCP scenario x RES scenario
    
    for r in range(0,NE): # RE scenario
        Path         = os.path.join(RECC_Paths.results_path,FolderlistB[r],'SysVar_TotalGHGFootprint.xls')
        Resultfile1  = xlrd.open_workbook(Path)
        Resultsheet1 = Resultfile1.sheet_by_name('Cover')
        UUID         = Resultsheet1.cell_value(3,2)
        Resultfile2  = xlrd.open_workbook(os.path.join(RECC_Paths.results_path,FolderlistB[r],'ODYM_RECC_ModelResults_' + UUID + '.xls'))
        Resultsheet2 = Resultfile2.sheet_by_name('Model_Results')
        for s in range(0,NS): # SSP scenario
            for c in range(0,NR):
                for t in range(0,45): # timeAnnEmsV_SecondarySteel[t,s,c,r] = Resultsheet2.cell_value(151+ 2*s +c,t+8)
                    AnnEmsV_PrimarySteel[t,s,c,r] = Resultsheet2.cell_value(19+ 2*s +c,t+8)
                    AnnEmsV_SecondarySteel[t,s,c,r] = Resultsheet2.cell_value(151+ 2*s +c,t+8)
                    
    Title      = ['primary_steel','secondary_steel']            
    Sector     = ['Residential_Buildings']
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
        
    
    return ASummaryV, AvgDecadalEmsV, MatSummaryV, AvgDecadalMatEmsV

# code for script to be run as standalone function
if __name__ == "__main__":
    main()
