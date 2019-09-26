# -*- coding: utf-8 -*-
"""
Created on Wed Oct 17 10:37:00 2018

@author: spauliuk
"""
def main(RegionalScope,PassVehList):
        
    import xlrd
    import numpy as np
    import matplotlib.pyplot as plt  
    import pylab
    import RECC_Paths # Import path file    
    import os
    
    #Sensitivity analysis folder order, by default, all strategies are off, one by one is implemented each at a time.
    #Baseline (no RE)
    #FabYieldImprovement
    #FabScrapDiversion
    #EoL_RR_Improvement
    #ChangeMaterialComposition
    #ReduceMaterialContent
    #ReUse_Materials
    #LifeTimeExtension
    #CarSharing
    #RideSharing
    #NoRecycling
    
        
    Region           = RegionalScope
    FolderlistV_Sens = PassVehList
    
    
    # Sensitivity plots
    
    NS = 3
    NR = 11
    
    CumEmsV_Sens         = np.zeros((NS,NR)) # SSP-Scenario x RES scenario
    AnnEmsV2030_Sens     = np.zeros((NS,NR)) # SSP-Scenario x RES scenario
    AnnEmsV2050_Sens     = np.zeros((NS,NR)) # SSP-Scenario x RES scenario
    AvgDecadalEms        = np.zeros((NS,NR,4)) # SSP-Scenario x RES scenario x 4 decades
    # for materials:
    MatCumEmsV_Sens      = np.zeros((NS,NR)) # SSP-Scenario x RES scenarioß
    MatAnnEmsV2030_Sens  = np.zeros((NS,NR)) # SSP-Scenario x RES scenario
    MatAnnEmsV2050_Sens  = np.zeros((NS,NR)) # SSP-Scenario x RES scenario
    MatAvgDecadalEms     = np.zeros((NS,NR,4)) # SSP-Scenario x RES scenario x 4 decades
    # for materials incl. recycling credit:
    MatCumEmsV_SensC     = np.zeros((NS,NR)) # SSP-Scenario x RES scenario
    MatAnnEmsV2030_SensC = np.zeros((NS,NR)) # SSP-Scenario x RES scenario
    MatAnnEmsV2050_SensC = np.zeros((NS,NR)) # SSP-Scenario x RES scenario
    MatAvgDecadalEmsC    = np.zeros((NS,NR,4)) # SSP-Scenario x RES scenario x 4 decades    
    
    for r in range(0,NR): # RE scenario
        
        Resultfile = xlrd.open_workbook(os.path.join(RECC_Paths.results_path,FolderlistV_Sens[r],'SysVar_TotalGHGFootprint.xls'))
        Resultsheet = Resultfile.sheet_by_name('TotalGHGFootprint')
        for s in range(0,NS): # SSP scenario
            for t in range(0,35): # time
                CumEmsV_Sens[s,r] += Resultsheet.cell_value(t +2, 2*(s+1))
            AnnEmsV2030_Sens[s,r]  = Resultsheet.cell_value(16  , 2*(s+1))
            AnnEmsV2050_Sens[s,r]  = Resultsheet.cell_value(36  , 2*(s+1))
            AvgDecadalEms[s,r,0]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(7,17)])/10
            AvgDecadalEms[s,r,1]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(17,27)])/10
            AvgDecadalEms[s,r,2]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(27,37)])/10
            AvgDecadalEms[s,r,3]   = sum([Resultsheet.cell_value(i, 2*(s+1)) for i in range(37,47)])/10
        # export material-related emissions
        Resultsheet1 = Resultfile.sheet_by_name('Cover')
        UUID         = Resultsheet1.cell_value(3,2)
        Resultfile2  = xlrd.open_workbook(os.path.join(RECC_Paths.results_path,FolderlistV_Sens[r],'ODYM_RECC_ModelResults_' + UUID + '.xls'))
        Resultsheet2 = Resultfile2.sheet_by_name('Model_Results')
        # Find the index for the recycling credit and others:
        rci = 1
        while True:
            if Resultsheet2.cell_value(rci, 0) == 'GHG emissions, recycling credits':
                break # that gives us the right index to read the recycling credit from the result table.
            rci += 1
        mci = 1
        while True:
            if Resultsheet2.cell_value(mci, 0) == 'GHG emissions, material cycle industries and their energy supply':
                break # that gives us the right index to read the recycling credit from the result table.
            mci += 1
            
        # Material results export
        for s in range(0,NS): # SSP scenario
            for t in range(0,35): # time until 2050 only!!! Cum. emissions until 2050.
                MatCumEmsV_Sens[s,r] += Resultsheet2.cell_value(mci+ 2*s +1,t+8)
            MatAnnEmsV2030_Sens[s,r]  = Resultsheet2.cell_value(mci+ 2*s +1,22)
            MatAnnEmsV2050_Sens[s,r]  = Resultsheet2.cell_value(mci+ 2*s +1,42)
            MatAvgDecadalEms[s,r,0]   = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for i in range(12,22)])/10
            MatAvgDecadalEms[s,r,1]   = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for i in range(22,32)])/10
            MatAvgDecadalEms[s,r,2]   = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for i in range(32,42)])/10
            MatAvgDecadalEms[s,r,3]   = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for i in range(42,52)])/10               
        # Material results export, including recycling credit
        for s in range(0,NS): # SSP scenario
            for t in range(0,35): # time until 2050 only!!! Cum. emissions until 2050.
                MatCumEmsV_SensC[s,r]+= Resultsheet2.cell_value(mci+ 2*s +1,t+8) + Resultsheet2.cell_value(rci+ 2*s +1,t+8)
            MatAnnEmsV2030_SensC[s,r] = Resultsheet2.cell_value(mci+ 2*s +1,22)  + Resultsheet2.cell_value(rci+ 2*s +1,22)
            MatAnnEmsV2050_SensC[s,r] = Resultsheet2.cell_value(mci+ 2*s +1,42)  + Resultsheet2.cell_value(rci+ 2*s +1,42)
            MatAvgDecadalEmsC[s,r,0]  = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for i in range(12,22)])/10 + sum([Resultsheet2.cell_value(rci+ 2*s +1,t) for i in range(12,22)])/10
            MatAvgDecadalEmsC[s,r,1]  = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for i in range(22,32)])/10 + sum([Resultsheet2.cell_value(rci+ 2*s +1,t) for i in range(22,32)])/10
            MatAvgDecadalEmsC[s,r,2]  = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for i in range(32,42)])/10 + sum([Resultsheet2.cell_value(rci+ 2*s +1,t) for i in range(32,42)])/10
            MatAvgDecadalEmsC[s,r,3]  = sum([Resultsheet2.cell_value(mci+ 2*s +1,t) for i in range(42,52)])/10 + sum([Resultsheet2.cell_value(rci+ 2*s +1,t) for i in range(42,52)])/10            
    
    ### Tornado plot for sensitivity
            
    MyColorCycle = pylab.cm.Set1(np.arange(0,1,0.1)) # select 12 colors from the 'Paired' color map.            
    
    Title = ['Passenger vehicles']
    Scens = ['LED','SSP1','SSP2']
    LWE   = ['Higher yield, manuf. efficiency','Fab scrap diversion','EoL_RR_Improvement','Material substitution','ReduceMaterialContent','ReUse_Materials','LifeTimeExtension','CarSharing','RideSharing','No recycling']
    
    #2030 emissions
    
    for m in range(0,NS): # SSP

        Data = AnnEmsV2030_Sens[:,1::]-np.einsum('S,n->Sn',AnnEmsV2030_Sens[:,0],np.ones(NR-1))
        Base = AnnEmsV2030_Sens[:,0]
            
        # plot results
    
        fig  = plt.figure(figsize=(5,NR))
        ax1  = plt.axes([0.08,0.08,0.85,0.9])
        
        Poss = np.arange(NR-1,0,-1)
        ax1.barh(Poss,Data[m,:], color=MyColorCycle, lw=0.4)

        # plot text and labels
        for mm in range(0,NR-1):
            plt.text(Data[m,:].min()*0.9, 9.9-mm, LWE[mm] + ': ' + ("%3.0f" % Data[m,mm]),fontsize=14,fontweight='bold')          
        
        plt.text(Data[m,:].min()*0.59, 10.5, 'Baseline: ' + ("%3.0f" % Base[m]) + ' Mt/yr.',fontsize=14,fontweight='bold')

        plt.title('2030 GHG emissions, sensitivity, ' + Region + '_ ' + Title[0] + '_' + Scens[m] + '.', fontsize = 18)
        plt.ylabel('RE strategies.', fontsize = 18)
        plt.xlabel('Mt of CO2-eq.', fontsize = 14)
        #plt.xticks([0.25,1.25,2.25,3.25,4.25,5.25])
        plt.yticks([])
        #ax1.set_xticklabels([], rotation =90, fontsize = 21, fontweight = 'normal')
        #plt_lgd  = plt.legend(handles = ProxyHandlesList,labels = LWE,shadow = False, prop={'size':16},ncol=1, loc = 'upper right' ,bbox_to_anchor=(1.91, 1)) 
        plt.axis([Data[m,:].min() -15, Data[m,:].max() +80, 0.3, 11.1])
    
        plt.show()
        fig_name = '2030_GHG_Sens_' + Region + '_ ' + Title[0] + '_' + Scens[m] + '_V2.png'
        fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = 400, bbox_inches='tight')             
      
        
    #2050 emissions
    
    for m in range(0,NS): # SSP
        
        Data = AnnEmsV2050_Sens[:,1::]-np.einsum('S,n->Sn',AnnEmsV2050_Sens[:,0],np.ones(NR-1))
        Base = AnnEmsV2050_Sens[:,0]
        
            
        # plot results

        fig  = plt.figure(figsize=(5,NR))
        ax1  = plt.axes([0.08,0.08,0.85,0.9])
            
        Poss = np.arange(NR-1,0,-1)
        ax1.barh(Poss,Data[m,:], color=MyColorCycle, lw=0.4)

        # plot text and labels
        for mm in range(0,NR-1):
            plt.text(Data[m,:].min()*0.9, 9.9-mm, LWE[mm] + ': ' + ("%3.0f" % Data[m,mm]),fontsize=14,fontweight='bold')          
        
        plt.text(Data[m,:].min()*0.59, 10.5, 'Baseline: ' + ("%3.0f" % Base[m]) + ' Mt/yr.',fontsize=14,fontweight='bold')

        plt.title('2050 GHG emissions, sensitivity, ' + Region + '_ ' + Title[0] + '_' + Scens[m] + '.', fontsize = 18)
        plt.ylabel('RE strategies.', fontsize = 18)
        plt.xlabel('Mt of CO2-eq.', fontsize = 14)
        #plt.xticks([0.25,1.25,2.25,3.25,4.25,5.25])
        plt.yticks([])
        #ax1.set_xticklabels([], rotation =90, fontsize = 21, fontweight = 'normal')
        #plt_lgd  = plt.legend(handles = ProxyHandlesList,labels = LWE,shadow = False, prop={'size':16},ncol=1, loc = 'upper right' ,bbox_to_anchor=(1.91, 1)) 
        plt.axis([Data[m,:].min() -15, Data[m,:].max() +80, 0.3, 11.1])
    
        plt.show()
        fig_name = '2050_GHG_Sens_' + Region + '_ ' + Title[0] + '_' + Scens[m] + '_V2.png'
        fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = 400, bbox_inches='tight')             
      

    
    #2050 cum. emissions
    
    for m in range(0,NS): # SSP
            Data = CumEmsV_Sens[:,1::]-np.einsum('S,n->Sn',CumEmsV_Sens[:,0],np.ones(NR-1))
            Base = CumEmsV_Sens[:,0]
        
            # plot results
        
            fig  = plt.figure(figsize=(5,NR))
            ax1  = plt.axes([0.08,0.08,0.85,0.9])
                
            Poss = np.arange(NR-1,0,-1)
    
            ax1.barh(Poss,Data[m,:], color=MyColorCycle, lw=0.4)
            
            # plot text and labels
            for mm in range(0,NR-1):
                plt.text(Data[m,:].min()  * 0.9, 9.9-mm, LWE[mm] + ': ' + ("%3.0f" % Data[m,mm]),fontsize=14,fontweight='bold')          
    
            plt.text(Data[m,:].min()*0.75, 10.5, 'Baseline: ' + ("%2.0f" % Base[m]) + ' Mt.',fontsize=14,fontweight='bold')
            plt.title('2016-2050 cum. GHG, sensitivity, ' + Region + '_ ' + Title[0] + '_' + Scens[m] + '.', fontsize = 18)
            plt.ylabel('RE strategies.', fontsize = 18)
            plt.xlabel('Mt of CO2-eq.', fontsize = 14)
            #plt.xticks([0.25,1.25,2.25,3.25,4.25,5.25])
            plt.yticks([])
            #ax1.set_xticklabels([], rotation =90, fontsize = 21, fontweight = 'normal')
            #plt_lgd  = plt.legend(handles = ProxyHandlesList,labels = LWE,shadow = False, prop={'size':16},ncol=1, loc = 'upper right' ,bbox_to_anchor=(1.91, 1)) 
            plt.axis([Data[m,:].min() *1.05, Data[m,:].max() *1.05, 0.3, 11.1])
        
            plt.show()
            fig_name = 'Cum_GHG_Sens_' + Region + '_ ' + Title[0] + '_' + Scens[m] + '.png'
            fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = 400, bbox_inches='tight')             
            
    return CumEmsV_Sens, AnnEmsV2030_Sens, AnnEmsV2050_Sens, AvgDecadalEms, MatCumEmsV_Sens, MatAnnEmsV2030_Sens, MatAnnEmsV2050_Sens, MatAvgDecadalEms, MatCumEmsV_SensC, MatAnnEmsV2030_SensC, MatAnnEmsV2050_SensC, MatAvgDecadalEmsC


# code for script to be run as standalone function
if __name__ == "__main__":
    main()
