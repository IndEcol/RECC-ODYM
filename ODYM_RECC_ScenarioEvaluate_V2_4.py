# -*- coding: utf-8 -*-
"""
Created on Fri Jun 14 05:18:48 2019

@author: spauliuk
"""

"""
File RECC_ScenarioEvaluate_V2_4.py

Script that runs the sensitivity and scnenario comparison scripts for different settings.

Section 1: single sector cascade
Section 2: multi-sector cascade
Section 3: Sensitivity plots
Section 4: Bar plot sufficiency

"""

# Import required libraries:
import os
import xlrd
import openpyxl
import numpy as np
import matplotlib.pyplot as plt 
import pylab

import RECC_Paths # Import path file

# The following SINGLE REGION scripts are called whenever there is a single cascade (for reb, pav, ...) or sensitivity analysis for a given region.
import ODYM_RECC_Cascade_V2_4
import ODYM_RECC_BarPlot_Eff_Suff_V2_4
import ODYM_RECC_Sensitivity_V2_4
import ODYM_RECC_Table_Extract_V2_4

# The following ALL REGION scripts are called when ALL 20 world regions are present in the result folder list.


# Define list of 20 regions, to be arranged 5 x 5, and corresponding data containers
RegionList20 = ['Global','G7','EU28','legend','legend', \
'R32USA','R32CAN','R32CHN','R32IND','R32JPN', \
'France','Germany','Italy','Poland','Spain', \
'UK','Oth_R32EU15','Oth_R32EU12-H','R32EU12-M','R5.2OECD_Other', \
'R5.2REF_Other','R5.2MNF_Other','R5.2SSA_Other','R5.2LAM_Other','R5.2ASIA_Other']

RegionList20Plot = ['Global','G7','EU28','legend','legend', \
'USA','CAN','CHN','IND','JPN', \
'France','Germany','Italy','Poland','Spain', \
'UK','Oth_EU15','Oth_EU12-H','EU12-M','OECD_Other', \
'REF_Other','MNF_Other','SSA_Other','LAM_Other','ASIA_Other']

PlotOrder      = [] # Will contain positions of countries/regions in 5x5 plot
TimeSeries_All = np.zeros((10,45,25,2,3,2)) # NX x Nt x Nr x NV x NS x NR / indicators x time x regions x sectors x SSP x RCP
# 0: system-wide GHG, no RES 1: system-wide GHG, all RES
# 2: material-related GHG, no RES, 3: material-related GHG, all RES,

PlotExpResolution = 200

# Number of scenarios:
NS = 3 # SSP
NR = 2 # RCP
    
###ScenarioSetting, sheet name of RECC_ModelConfig_List.xlsx to be selected:
ScenarioSetting = 'Evaluate_pav_reb_Cascade' # run eval and plot scripts for selected regions and sectors only
#ScenarioSetting = 'Evaluate_pav_reb_Cascade_all' # run eval and plot scripts for all regions and sectors
#ScenarioSetting = 'Germany_detail_evaluate' # run eval and plot scripts for Germany case study only
#ScenarioSetting = 'Global_all_evaluate' # run eval and plot scripts for all regions and all sectors
#ScenarioSetting = 'Evaluate_TestRun' # Test run evaluate


# open scenario sheet
ModelConfigListFile  = xlrd.open_workbook(os.path.join(RECC_Paths.recc_path,'RECC_ModelConfig_List_V2_4.xlsx'))
ModelEvalListSheet   = ModelConfigListFile.sheet_by_name(ScenarioSetting)

# open result summary file
mywb  = openpyxl.load_workbook(os.path.join(RECC_Paths.results_path,'RECC_Global_Results_Template_CascSens.xlsx')) # for total emissions
mywb4 = openpyxl.load_workbook(os.path.join(RECC_Paths.results_path,'RECC_Global_Results_Template_Overview.xlsx')) # for emissions to be reported in Tables.

#Read control lines and execute main model script
Row   = 1

Table_Annual  = np.zeros((3,8,NS,NR)) # 2050 annual system emissions, cascade steps x SSP scenarios x RCP scenarios.
Table_CumEms  = np.zeros((3,8,NS,NR)) # 2016-2050 (!) cumulative system emissions, cascade steps x SSP scenarios x RCP scenarios.
MatStocksTab1 = np.zeros((9,6)) # Material stocks for table, LED.
MatStocksTab2 = np.zeros((9,6)) # Material stocks for table, SSP1.
MatStocksTab3 = np.zeros((9,6)) # Material stocks for table, SSP2.

CascadeFlag1  = False
CascadeFlag2  = False
SensitiFlag1  = False

SingleSectList = [] # For model runs not part of sensitivity or cascade, used for efficiency-sufficiency bar plot
# search for script config list entry
while ModelEvalListSheet.cell_value(Row, 1)  != 'ENDOFLIST':
    if ModelEvalListSheet.cell_value(Row, 1) != '':
        FolderList    = []
        
        MultiSectorList= []
        RegionalScope  = ModelEvalListSheet.cell_value(Row, 1)
        Setting        = ModelEvalListSheet.cell_value(Row, 2) # cascade or sensitivity
        print(RegionalScope)
        
    if Setting == 'Cascade_pav':
        CascadeFlag1     = True
        SectorString     = 'pav'
        Vsheet           = mywb[RegionalScope  + '_Vehicles']
        NE               = 7 # 7 for vehs. and 6 for buildings
        
    if Setting == 'Cascade_reb':        
        CascadeFlag1     = True
        SectorString     = 'reb'
        Vsheet           = mywb[RegionalScope  + '_ResBuildings']
        NE               = 6 # 7 for vehs. and 6 for buildings     

    if Setting == 'Cascade_nrb':        
        CascadeFlag1     = True
        SectorString     = 'nrb'
        Vsheet           = mywb[RegionalScope  + '_NonResBuildings']
        NE               = 6 # 7 for vehs. and 6 for buildings     
        
    if Setting == 'Cascade_pav_reb':
        CascadeFlag2     = True
        SectorString     = 'pav_reb'
        NE               = 8 # 8 for vehs, res and nonres buildings
            
    if Setting == 'Cascade_pav_reb_nrb':
        CascadeFlag2     = True
        SectorString     = 'pav_reb_nrb'
        NE               = 8 # 8 for vehs, res and nonres buildings
        
    CascCols = [5,13]        
        
    if CascadeFlag1 is True: #extract results for this cascade and store
        CascadeFlag1 = False
        print('Cascade_' + RegionalScope + '_' + SectorString)
        for m in range(0,NE):
            FolderList.append(ModelEvalListSheet.cell_value(Row +m, 3))
        # run the cascade plot function
        ASummary, AvgDecadalEms, MatSummary, AvgDecadalMatEms, RecCredit, UsePhaseSummary, ManSummary, ForSummary, AvgDecadalUseEms, AvgDecadalManEms, AvgDecadalForEms, AvgDecadalRecEms, CumEms, AnnEms2050, MatStocks, TimeSeries_R = ODYM_RECC_Cascade_V2_4.main(RegionalScope,FolderList,SectorString)
        # write results summary to Excel
        for R in range(0,NR):
            for r in range(0,3):
                for c in range(0,NE):
                    Vsheet.cell(row = r+3,  column   = c +CascCols[R]).value   = ASummary[r,R,c]       
                    Vsheet.cell(row = r+9,  column   = c +CascCols[R]).value   = ASummary[r+3,R,c]       
                    Vsheet.cell(row = r+15, column   = c +CascCols[R]).value   = ASummary[r+6,R,c]
                    Vsheet.cell(row = r+36, column   = c +CascCols[R]).value   = ASummary[r+9,R,c]
                    for d in range(0,4):
                        Vsheet.cell(row = d*3 + r + 21,column = c +CascCols[R]).value  = AvgDecadalEms[r,R,c,d]
            for r in range(0,3):
                for c in range(0,NE):
                    Vsheet.cell(row = r+45,  column  = c +CascCols[R]).value  = UsePhaseSummary[r,R,c]       
                    Vsheet.cell(row = r+51,  column  = c +CascCols[R]).value  = UsePhaseSummary[r+3,R,c]       
                    Vsheet.cell(row = r+57,  column  = c +CascCols[R]).value  = UsePhaseSummary[r+6,R,c]
                    Vsheet.cell(row = r+78,  column  = c +CascCols[R]).value  = UsePhaseSummary[r+9,R,c]
                    for d in range(0,4):
                        Vsheet.cell(row = d*3 + r + 63,column = c +CascCols[R]).value  = AvgDecadalUseEms[r,R,c,d]                    
            for r in range(0,3):
                for c in range(0,NE):
                    Vsheet.cell(row = r+87,  column  = c +CascCols[R]).value  = MatSummary[r,R,c]       
                    Vsheet.cell(row = r+93,  column  = c +CascCols[R]).value  = MatSummary[r+3,R,c]       
                    Vsheet.cell(row = r+99,  column  = c +CascCols[R]).value  = MatSummary[r+6,R,c]
                    Vsheet.cell(row = r+120, column  = c +CascCols[R]).value  = MatSummary[r+9,R,c]
                    for d in range(0,4):
                        Vsheet.cell(row = d*3 + r + 105,column = c +CascCols[R]).value  = AvgDecadalMatEms[r,R,c,d] 
            for r in range(0,3):
                for c in range(0,NE):
                    Vsheet.cell(row = r+129,  column = c +CascCols[R]).value  = ManSummary[r,R,c]       
                    Vsheet.cell(row = r+135,  column = c +CascCols[R]).value  = ManSummary[r+3,R,c]       
                    Vsheet.cell(row = r+141,  column = c +CascCols[R]).value  = ManSummary[r+6,R,c]
                    Vsheet.cell(row = r+162,  column = c +CascCols[R]).value  = ManSummary[r+9,R,c]
                    for d in range(0,4):
                        Vsheet.cell(row = d*3 + r + 147,column = c +CascCols[R]).value  = AvgDecadalManEms[r,R,c,d]                         
            for r in range(0,3):
                for c in range(0,NE):
                    Vsheet.cell(row = r+171, column  = c +CascCols[R]).value  = ForSummary[r,R,c]       
                    Vsheet.cell(row = r+177, column  = c +CascCols[R]).value  = ForSummary[r+3,R,c]       
                    Vsheet.cell(row = r+183, column  = c +CascCols[R]).value  = ForSummary[r+6,R,c]
                    Vsheet.cell(row = r+204, column  = c +CascCols[R]).value  = ForSummary[r+9,R,c]
                    for d in range(0,4):
                        Vsheet.cell(row = d*3 + r + 189,column = c +CascCols[R]).value  = AvgDecadalForEms[r,R,c,d]                         
            for r in range(0,3):
                for c in range(0,NE):
                    Vsheet.cell(row = r+213,  column = c +CascCols[R]).value  = RecCredit[r,R,c]       
                    Vsheet.cell(row = r+219,  column = c +CascCols[R]).value  = RecCredit[r+3,R,c]       
                    Vsheet.cell(row = r+225,  column = c +CascCols[R]).value  = RecCredit[r+6,R,c]
                    Vsheet.cell(row = r+246,  column = c +CascCols[R]).value  = RecCredit[r+9,R,c]
                    for d in range(0,4):
                        Vsheet.cell(row = d*3 + r + 231,column = c +CascCols[R]).value  = AvgDecadalRecEms[r,R,c,d]   
        
        # Store results in time series array
        if SectorString == 'pav' or SectorString == 'reb':
            if SectorString == 'pav':
                SectorIndex = 0
            if SectorString == 'reb':
                SectorIndex = 1
            RegPos = RegionList20.index(RegionalScope)
            PlotOrder.append(RegPos)
            TimeSeries_All[0,:,RegPos,SectorIndex,:,:] = TimeSeries_R[0,0,:,:,:]  # system-wide GHG, no RES
            TimeSeries_All[1,:,RegPos,SectorIndex,:,:] = TimeSeries_R[0,-1,:,:,:] # system-wide GHG, full RES
            TimeSeries_All[2,:,RegPos,SectorIndex,:,:] = TimeSeries_R[1,0,:,:,:]  # matcycle GHG, no RES
            TimeSeries_All[3,:,RegPos,SectorIndex,:,:] = TimeSeries_R[1,-1,:,:,:] # matcycle GHG, full RES
            TimeSeries_All[4,:,RegPos,SectorIndex,:,:] = TimeSeries_R[2,0,:,:,:]  # primary production total, no RES
            TimeSeries_All[5,:,RegPos,SectorIndex,:,:] = TimeSeries_R[2,-1,:,:,:] # primary production total, full RES
            TimeSeries_All[6,:,RegPos,SectorIndex,:,:] = TimeSeries_R[3,0,:,:,:]  # secondary production total, no RES
            TimeSeries_All[7,:,RegPos,SectorIndex,:,:] = TimeSeries_R[3,-1,:,:,:] # secondary production total, full RES
            
        RCP_Matstocks = 1 # MatStocks are plotted for RCP2.6 only
        
        if Setting == 'Cascade_pav':                    
            # store other results
            Table_Annual[0,0:-1,1,:]= AnnEms2050[1,:,:].transpose().copy()
            Table_CumEms[0,0:-1,1,:]= CumEms[1,:,:].transpose().copy()
            MatStocksTab1[0,:]      = MatStocks[4,:,0,RCP_Matstocks,0].copy()
            MatStocksTab1[1,:]      = MatStocks[34,:,0,RCP_Matstocks,0].copy()
            MatStocksTab1[2,:]      = MatStocks[34,:,0,RCP_Matstocks,-1].copy()
            MatStocksTab2[0,:]      = MatStocks[4,:,1,RCP_Matstocks,0].copy()
            MatStocksTab2[1,:]      = MatStocks[34,:,1,RCP_Matstocks,0].copy()
            MatStocksTab2[2,:]      = MatStocks[34,:,1,RCP_Matstocks,-1].copy()
            MatStocksTab3[0,:]      = MatStocks[4,:,2,RCP_Matstocks,0].copy()
            MatStocksTab3[1,:]      = MatStocks[34,:,2,RCP_Matstocks,0].copy()
            MatStocksTab3[2,:]      = MatStocks[34,:,2,RCP_Matstocks,-1].copy()
                    
        if Setting == 'Cascade_reb':
            # store other results
            Table_Annual[1,0:5,1,:] = AnnEms2050[1,:,0:-1].transpose().copy()
            Table_Annual[1,7,1,:]   = AnnEms2050[1,:,-1].copy()
            Table_CumEms[1,0:5,1,:] = CumEms[1,:,0:-1].transpose().copy()                    
            Table_CumEms[1,7,1,:]   = CumEms[1,:,-1].copy()        
            MatStocksTab1[3,:]      = MatStocks[4,:,0,RCP_Matstocks,0].copy()
            MatStocksTab1[4,:]      = MatStocks[34,:,0,RCP_Matstocks,0].copy()
            MatStocksTab1[5,:]      = MatStocks[34,:,0,RCP_Matstocks,-1].copy()            
            MatStocksTab2[3,:]      = MatStocks[4,:,1,RCP_Matstocks,0].copy()
            MatStocksTab2[4,:]      = MatStocks[34,:,1,RCP_Matstocks,0].copy()
            MatStocksTab2[5,:]      = MatStocks[34,:,1,RCP_Matstocks,-1].copy()            
            MatStocksTab3[3,:]      = MatStocks[4,:,2,RCP_Matstocks,0].copy()
            MatStocksTab3[4,:]      = MatStocks[34,:,2,RCP_Matstocks,0].copy()
            MatStocksTab3[5,:]      = MatStocks[34,:,2,RCP_Matstocks,-1].copy()            
    
        if Setting == 'Cascade_nrb':                        
            # store other results
            Table_Annual[1,0:5,2,:] = AnnEms2050[1,:,0:-1].transpose().copy()
            Table_Annual[1,7,2,:]   = AnnEms2050[1,:,-1].copy()
            Table_CumEms[1,0:5,2,:] = CumEms[1,:,0:-1].transpose().copy()                    
            Table_CumEms[1,7,2,:]   = CumEms[1,:,-1].copy()      
            MatStocksTab1[6,:]      = MatStocks[4,:,0,RCP_Matstocks,0].copy()
            MatStocksTab1[7,:]      = MatStocks[34,:,0,RCP_Matstocks,0].copy()
            MatStocksTab1[8,:]      = MatStocks[34,:,0,RCP_Matstocks,-1].copy()                 
            MatStocksTab2[6,:]      = MatStocks[4,:,1,RCP_Matstocks,0].copy()
            MatStocksTab2[7,:]      = MatStocks[34,:,1,RCP_Matstocks,0].copy()
            MatStocksTab2[8,:]      = MatStocks[34,:,1,RCP_Matstocks,-1].copy()                 
            MatStocksTab3[6,:]      = MatStocks[4,:,2,RCP_Matstocks,0].copy()
            MatStocksTab3[7,:]      = MatStocks[34,:,2,RCP_Matstocks,0].copy()
            MatStocksTab3[8,:]      = MatStocks[34,:,2,RCP_Matstocks,-1].copy()                 
                    
    if CascadeFlag2 is True: #extract results for this cascade and store    
        CascadeFlag2 = False
        for m in range(0,NE):
            MultiSectorList.append(ModelEvalListSheet.cell_value(Row +m, 3))  
        GHG_TableX = ODYM_RECC_Table_Extract_V2_4.main(RegionalScope,MultiSectorList)        
        # write results summary as Table 2 to Excel
        Gsheet = mywb4['GHG_Overview']
        print('GHG_Overview_' + RegionalScope)
        for r in range(0,4):
            for c in range(0,6):
                for R in range(0,2):
                    Gsheet.cell(row = r+4 + 8*R, column = c+4).value  = GHG_TableX[r,c,R]        

        # run the cascade plots for the three sectors
        ASummary, AvgDecadalEms, MatSummary, AvgDecadalMatEms, RecCredit, AvgDecadalMatEmsC, CumEms, AnnEms2050, MatStocks, TimeSeries_R = ODYM_RECC_Cascade_V2_4.main(RegionalScope,MultiSectorList,SectorString)
                    
        # plot GHG overview figure global
        if RegionalScope == 'Global':
            ColOrder     = [11,4,0,18,8,16,2,6,15]
            Xoff         = [1,1.7,2.7,3.4,4.4,5.1]
            MyColorCycle = pylab.cm.tab20(np.arange(0,1,0.05)) # select 20 colors from the 'tab20' color map. 
            Data_Cum_Abs = CumEms
            Data_Cum_pc  = Data_Cum_Abs / np.einsum('SR,E->SRE',Data_Cum_Abs[:,:,0],np.ones(8))
            Data_50_Abs  = np.einsum('ESR->SRE',TimeSeries_R[0,:,35,:,:])
            Data_50_pc   = Data_50_Abs / np.einsum('SR,E->SRE',Data_50_Abs[:,:,0],np.ones(8))
            Cum_savings  = Data_Cum_Abs[:,:,0] - Data_Cum_Abs[:,:,-1]
            Ann_savings  = Data_50_Abs[:,:,0]  - Data_50_Abs[:,:,-1]
            LWE          = ['higher yields', 're-use/longer use','material subst.','down-sizing','car-sharing','ride-sharing','More intense bld. use','Bottom line']
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
            plt.plot([-1,8],[1.1,1.1],linestyle = '-', linewidth = 0.5, color = 'k')
            plt.plot([2.45,2.45],[-1,3],linestyle = '--', linewidth = 0.5, color = 'k')
            plt.plot([4.15,4.15],[-1,3],linestyle = '--', linewidth = 0.5, color = 'k')
            for fca in range(0,NE):
                ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[ColOrder[fca],:])) # create proxy artist for legend

            # plot emissions reductions
            for mS in range(0,3):
                for mR in range(0,2): 
                    plt.text(Xoff[mS*2+mR]+0.2, 1.9, ("%3.0f" % (Cum_savings[mS,mR]/1000)) + ' Gt',fontsize=16, rotation=90, fontweight='normal')
                    plt.text(Xoff[mS*2+mR]+0.2, 0.5, ("%3.0f" % Ann_savings[mS,mR]) + ' Mt',fontsize=16, rotation=90, fontweight='normal')
                
            # plot text and labels
            plt.text(0.85, 0.97, '2050 annual emissions',fontsize=11, rotation=90, fontweight='bold')          
            plt.text(0.85, 2.2, '2016-50 cum. emissions',fontsize=11, rotation=90, fontweight='bold')          

            plt.title('RE strats. and GHG emissions, global, pav+reb.', fontsize = 18)
            
            plt.ylabel('GHG reductions, system-wide, %.', fontsize = 18)
            plt.xticks(XTicks)
            plt.yticks(YTicks, fontsize =12)
            ax1.set_xticklabels(['LED+NoPol','LED+RCP2.6','SSP1+NoPol','SSP1+RCP2.6','SSP2+NoPol','SSP2+RCP2.6'], rotation =90, fontsize = 15, fontweight = 'normal')
            ax1.set_yticklabels(YTickLabels, rotation =0, fontsize = 15, fontweight = 'normal')
            plt.legend(handles = reversed(ProxyHandlesList),labels = LWE,shadow = False, prop={'size':12},ncol=1, loc = 'upper right' ,bbox_to_anchor=(1.40, 1)) 
            #plt.axis([-0.2, 7.7, 0.9*Right, 1.02*Left])
            plt.axis([0.7, 5.8, -0.1, 2.3])
        
            plt.show()
            fig_name = 'Overview_GHG_Global.png'
            fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = PlotExpResolution, bbox_inches='tight')   
                
                
    if Setting == 'Sensitivity_pav':
        SensitiFlag1     = True
        SectorString     = 'pav'
        NE               = 11 # 11 for vehs. and 10 for buildings 
        SensCols         = [6,18] 
        
    if Setting == 'Sensitivity_reb':        
        SensitiFlag1     = True
        SectorString     = 'reb'
        NE               = 10 # 11 for vehs. and 10 for buildings        
        SensCols         = [35,47] 
        
    if Setting == 'Sensitivity_nrb':        
        SensitiFlag1     = True
        SectorString     = 'nrb'
        NE               = 10 # 11 for vehs. and 10 for buildings        
        SensCols         = [63,75] 
    
    SensRows             = [4,9,14,19,24,40,45,50,55,60,76,81,86,91,96,112,117,122,127,132,148,153,158,163,168,184,189,194,199,204]
        
    if SensitiFlag1 is True:
        SensitiFlag1 = False
        for m in range(0,int(NE)):
            FolderList.append(ModelEvalListSheet.cell_value(Row +m, 3))
        # run the ODYM-RECC sensitivity analysis for pav
        CumEms_Sens2050, CumEms_Sens2060, AnnEms2030_Sens, AnnEms2050_Sens, AvgDecadalEms, UseCumEms2050, UseCumEms2060, UseAnnEms2030, UseAnnEms2050, AvgDecadalUseEms, MatCumEms2050, MatCumEms2060, MatAnnEms2030, MatAnnEms2050, AvgDecadalMatEms, ManCumEms2050, ManCumEms2060, ManAnnEms2030, ManAnnEms2050, AvgDecadalManEms, ForCumEms2050, ForCumEms2060, ForAnnEms2030, ForAnnEms2050, AvgDecadalForEms, RecCreditCum2050, RecCreditCum2060, RecCreditAnn2030, RecCreditAnn2050, RecCreditAvgDec = ODYM_RECC_Sensitivity_V2_4.main(RegionalScope,FolderList,SectorString)        

        # write results summary to Excel
        Ssheet  = mywb['Sensitivity_'  + RegionalScope]
        print('Sensitivity_' + RegionalScope + '_' + SectorString)
        for R in range(0,NR):
            for r in range(0,3):
                for c in range(0,NE):
                    Ssheet.cell(row = r +SensRows[0], column   = c +SensCols[R]).value   = AnnEms2030_Sens[r,R,c]
                    Ssheet.cell(row = r +SensRows[1], column   = c +SensCols[R]).value   = AnnEms2050_Sens[r,R,c]
                    Ssheet.cell(row = r +SensRows[2], column   = c +SensCols[R]).value   = CumEms_Sens2050[r,R,c]
                    Ssheet.cell(row = r +SensRows[3], column   = c +SensCols[R]).value   = CumEms_Sens2060[r,R,c]
                    for d in range(0,4):
                        Ssheet.cell(row = d*3 + r + SensRows[4],column  = c +SensCols[R]).value   = AvgDecadalEms[r,R,c,d]
            
            for r in range(0,3):
                for c in range(0,NE):
                    Ssheet.cell(row = r +SensRows[5], column  = c +SensCols[R]).value    = UseAnnEms2030[r,R,c]
                    Ssheet.cell(row = r +SensRows[6], column  = c +SensCols[R]).value    = UseAnnEms2050[r,R,c]
                    Ssheet.cell(row = r +SensRows[7], column  = c +SensCols[R]).value    = UseCumEms2050[r,R,c]
                    Ssheet.cell(row = r +SensRows[8], column  = c +SensCols[R]).value    = UseCumEms2060[r,R,c]
                    for d in range(0,4):
                        Ssheet.cell(row = d*3 + r + SensRows[9],column = c +SensCols[R]).value   = AvgDecadalUseEms[r,R,c,d]       
                       
            for r in range(0,3):
                for c in range(0,NE):
                    Ssheet.cell(row = r +SensRows[10], column  = c +SensCols[R]).value   = MatAnnEms2030[r,R,c]
                    Ssheet.cell(row = r +SensRows[11], column  = c +SensCols[R]).value   = MatAnnEms2050[r,R,c]
                    Ssheet.cell(row = r +SensRows[12], column  = c +SensCols[R]).value   = MatCumEms2050[r,R,c]
                    Ssheet.cell(row = r +SensRows[13], column  = c +SensCols[R]).value   = MatCumEms2060[r,R,c]
                    for d in range(0,4):
                        Ssheet.cell(row = d*3 + r + SensRows[14],column = c +SensCols[R]).value   = AvgDecadalMatEms[r,R,c,d] 

            for r in range(0,3):
                for c in range(0,NE):
                    Ssheet.cell(row = r +SensRows[15], column  = c +SensCols[R]).value   = ManAnnEms2030[r,R,c]
                    Ssheet.cell(row = r +SensRows[16], column  = c +SensCols[R]).value   = ManAnnEms2050[r,R,c]
                    Ssheet.cell(row = r +SensRows[17], column  = c +SensCols[R]).value   = ManCumEms2050[r,R,c]
                    Ssheet.cell(row = r +SensRows[18], column  = c +SensCols[R]).value   = ManCumEms2060[r,R,c]
                    for d in range(0,4):
                        Ssheet.cell(row = d*3 + r + SensRows[19],column = c +SensCols[R]).value   = AvgDecadalManEms[r,R,c,d] 

            for r in range(0,3):
                for c in range(0,NE):
                    Ssheet.cell(row = r +SensRows[20], column  = c +SensCols[R]).value   = ForAnnEms2030[r,R,c]
                    Ssheet.cell(row = r +SensRows[21], column  = c +SensCols[R]).value   = ForAnnEms2050[r,R,c]
                    Ssheet.cell(row = r +SensRows[22], column  = c +SensCols[R]).value   = ForCumEms2050[r,R,c]
                    Ssheet.cell(row = r +SensRows[23], column  = c +SensCols[R]).value   = ForCumEms2060[r,R,c]
                    for d in range(0,4):
                        Ssheet.cell(row = d*3 + r + SensRows[24],column = c +SensCols[R]).value   = AvgDecadalForEms[r,R,c,d] 

            for r in range(0,3):
                for c in range(0,NE):
                    Ssheet.cell(row = r +SensRows[25], column  = c +SensCols[R]).value   = RecCreditAnn2030[r,R,c]
                    Ssheet.cell(row = r +SensRows[26], column  = c +SensCols[R]).value   = RecCreditAnn2050[r,R,c]
                    Ssheet.cell(row = r +SensRows[27], column  = c +SensCols[R]).value   = RecCreditCum2050[r,R,c]
                    Ssheet.cell(row = r +SensRows[28], column  = c +SensCols[R]).value   = RecCreditCum2060[r,R,c]
                    for d in range(0,4):
                        Ssheet.cell(row = d*3 + r + SensRows[29],column = c +SensCols[R]).value   = RecCreditAvgDec[r,R,c,d]                         
                        
    if Setting == 'Do_not_include':
        NE = 1
        SingleSectList.append(ModelEvalListSheet.cell_value(Row, 3))

    # forward counter   
    Row += NE
        
# run the efficieny_sufficieny plots, only if 4 single sectors in result list
if len(SingleSectList) == 4:
    ODYM_RECC_BarPlot_Eff_Suff_V2_4.main(RegionalScope,MultiSectorList,SingleSectList)      

# store overview tables
# Done for each region, overwritten each time, data for LAST REGION remain.
WFsheet = mywb4['CascadeBySector']
v = 1 # SSP1
for u in range(0,8):
    for c in range(0,2):
        for z in range(0,3):
            WFsheet.cell(row = u+4,  column = z+3 +4*c).value  = Table_Annual[z,u,v,c]     
            WFsheet.cell(row = u+14, column = z+3 +4*c).value  = Table_CumEms[z,u,v,c]       
            
WFsheet = mywb4['MatStocksBySector']
for u in range(0,9):
    for v in range(0,6):
        WFsheet.cell(row = u+3 , column = v+6).value  = MatStocksTab1[u,v]     
        WFsheet.cell(row = u+14, column = v+6).value  = MatStocksTab2[u,v]     
        WFsheet.cell(row = u+25, column = v+6).value  = MatStocksTab3[u,v]     
    
mywb.save(os.path.join(RECC_Paths.results_path, 'RECC_Global_Results_SystemGHG_V2_4.xlsx'))      
mywb4.save(os.path.join(RECC_Paths.results_path,'RECC_Global_Results_Tables_V2_4.xlsx'))      
   
# plot SSP1 time series in 5x5 plot:
# TimeSeries_All indices: NX x Nt x Nr x NV x NS x NR / indicators x time x regions x sectors x SSP x RCP
# 0: system-wide GHG, no RES 1: system-wide GHG, all RES
# 2: material-related GHG, no RES, 3: material-related GHG, all RES,

# System-wide GHG
MyColorCycle = pylab.cm.Set1(np.arange(0,1,0.1)) # select 12 colors from the 'Paired' color map. 
plt.rcParams['axes.labelsize'] = 7
fig_names    = ['GHG_SSP1_pav_5x5.png','GHG_SSP1_reb_5x5.png']
fig_titles   = ['System-wide GHG, SSP1, pav, Mt/yr','System-wide GHG, SSP1, reb, Mt/yr']
LegendLables = ['RCP NoPol, full RES','RCP 2.6, full RES','RCP NoPol, no RES','RCP 2.6, no RES']
for Sect in range(0,2):
    fig, axs = plt.subplots(5, 5, sharex=True, gridspec_kw={'hspace': 0.6, 'wspace': 0.4})
    for plotNo in PlotOrder:
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[0,:,plotNo,Sect,1,0],color=MyColorCycle[4,:], lw=0.6, linestyle='--') # System GHG, sector, Baseline, no RES
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[0,:,plotNo,Sect,1,1],color=MyColorCycle[7,:], lw=0.6, linestyle='--') # System GHG, sector, RCP2.6, no RES
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[1,:,plotNo,Sect,1,0],color=MyColorCycle[4,:], lw=0.8, linestyle='-') # System GHG, sector, Baseline, full RES
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[1,:,plotNo,Sect,1,1],color=MyColorCycle[7,:], lw=0.8, linestyle='-') # System GHG, sector, RCP2.6, full RES
        axs[plotNo//5, plotNo%5].set_ylim(ymin=0)
        axs[plotNo//5, plotNo%5].set_title(RegionList20Plot[plotNo], fontsize=7)
        #axs[plotNo//5, plotNo%5].set_yticklabels(fontsize = 6)
        axs[plotNo//5, plotNo%5].tick_params(axis='x', labelsize=6)
        axs[plotNo//5, plotNo%5].tick_params(axis='y', labelsize=6)
    # legend plot:
    axs[0,4].plot(2016,0,color=MyColorCycle[4,:], lw=0.8, linestyle='-') # System GHG, sector, Baseline, full RES
    axs[0,4].plot(2016,0,color=MyColorCycle[7,:], lw=0.8, linestyle='-') # System GHG, sector, RCP2.6, full RES
    axs[0,4].plot(2016,0,color=MyColorCycle[4,:], lw=0.6, linestyle='--') # System GHG, sector, Baseline, no RES
    axs[0,4].plot(2016,0,color=MyColorCycle[7,:], lw=0.6, linestyle='--') # System GHG, sector, RCP2.6, no RES
    axs[0,4].legend(LegendLables,shadow = False, prop={'size':5}, loc = 'upper right')    
    axs[0,3].axis('off')
    axs[0,4].axis('off')
    fig.suptitle(fig_titles[Sect], fontsize=14)
    for xm in range(0,5):
        plt.sca(axs[4,xm])
        plt.xticks([2020,2030,2040,2050,2060], ['2020','2030','2040','2050','2060'], rotation =90, fontsize = 6, fontweight = 'normal')
    plt.show()
    fig_name = fig_names[Sect]
    fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = PlotExpResolution, bbox_inches='tight')  
    
# Material cycle GHG
MyColorCycle = pylab.cm.Set1(np.arange(0,1,0.1)) # select 12 colors from the 'Paired' color map. 
plt.rcParams['axes.labelsize'] = 7
fig_names    = ['GHGMat_SSP1_pav_5x5.png','GHGMat_SSP1_reb_5x5.png']
fig_titles   = ['Matcycle GHG, SSP1, pav, Mt/yr','Matcycle GHG, SSP1, reb, Mt/yr']
LegendLables = ['RCP NoPol, full RES','RCP 2.6, full RES','RCP NoPol, no RES','RCP 2.6, no RES']
for Sect in range(0,2):
    fig, axs = plt.subplots(5, 5, sharex=True, gridspec_kw={'hspace': 0.6, 'wspace': 0.4})
    for plotNo in PlotOrder:
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[2,:,plotNo,Sect,1,0],color=MyColorCycle[4,:], lw=0.6, linestyle='--') # System GHG, sector, Baseline, no RES
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[2,:,plotNo,Sect,1,1],color=MyColorCycle[7,:], lw=0.6, linestyle='--') # System GHG, sector, RCP2.6, no RES
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[3,:,plotNo,Sect,1,0],color=MyColorCycle[4,:], lw=0.8, linestyle='-') # System GHG, sector, Baseline, full RES
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[3,:,plotNo,Sect,1,1],color=MyColorCycle[7,:], lw=0.8, linestyle='-') # System GHG, sector, RCP2.6, full RES
        axs[plotNo//5, plotNo%5].set_ylim(ymin=0)
        axs[plotNo//5, plotNo%5].set_title(RegionList20Plot[plotNo], fontsize=7)
        #axs[plotNo//5, plotNo%5].set_yticklabels(fontsize = 6)
        axs[plotNo//5, plotNo%5].tick_params(axis='x', labelsize=6)
        axs[plotNo//5, plotNo%5].tick_params(axis='y', labelsize=6)
    # legend plot:
    axs[0,4].plot(2016,0,color=MyColorCycle[4,:], lw=0.8, linestyle='-') # System GHG, sector, Baseline, full RES
    axs[0,4].plot(2016,0,color=MyColorCycle[7,:], lw=0.8, linestyle='-') # System GHG, sector, RCP2.6, full RES
    axs[0,4].plot(2016,0,color=MyColorCycle[4,:], lw=0.6, linestyle='--') # System GHG, sector, Baseline, no RES
    axs[0,4].plot(2016,0,color=MyColorCycle[7,:], lw=0.6, linestyle='--') # System GHG, sector, RCP2.6, no RES
    axs[0,4].legend(LegendLables,shadow = False, prop={'size':5}, loc = 'upper right')         
    axs[0,3].axis('off')
    axs[0,4].axis('off')
    fig.suptitle(fig_titles[Sect], fontsize=14)
    for xm in range(0,5):
        plt.sca(axs[4,xm])
        plt.xticks([2020,2030,2040,2050,2060], ['2020','2030','2040','2050','2060'], rotation =90, fontsize = 6, fontweight = 'normal')
    plt.show()
    fig_name = fig_names[Sect]
    fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = PlotExpResolution, bbox_inches='tight')    

# Total secondary material production
MyColorCycle = pylab.cm.Set1(np.arange(0,1,0.1)) # select 12 colors from the 'Paired' color map. 
plt.rcParams['axes.labelsize'] = 7
fig_names    = ['SecMat_5x5_pav.png','SecMat_5x5_reb.png']
fig_titles   = ['Total secondary material, SSP1, pav, Mt/yr','Total secondary material, SSP1, reb, Mt/yr']
LegendLables = ['RCP NoPol, full RES','RCP 2.6, full RES','RCP NoPol, no RES','RCP 2.6, no RES']
for Sect in range(0,2):
    fig, axs = plt.subplots(5, 5, sharex=True, gridspec_kw={'hspace': 0.6, 'wspace': 0.4})
    for plotNo in PlotOrder:
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[6,:,plotNo,Sect,1,0],color=MyColorCycle[4,:], lw=0.6, linestyle='--') # System GHG, sector, Baseline, no RES
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[6,:,plotNo,Sect,1,1],color=MyColorCycle[7,:], lw=0.6, linestyle='--') # System GHG, sector, RCP2.6, no RES
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[7,:,plotNo,Sect,1,0],color=MyColorCycle[4,:], lw=0.8, linestyle='-') # System GHG, sector, Baseline, full RES
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[7,:,plotNo,Sect,1,1],color=MyColorCycle[7,:], lw=0.8, linestyle='-') # System GHG, sector, RCP2.6, full RES
        axs[plotNo//5, plotNo%5].set_ylim(ymin=0)
        axs[plotNo//5, plotNo%5].set_title(RegionList20Plot[plotNo], fontsize=7)
        #axs[plotNo//5, plotNo%5].set_yticklabels(fontsize = 6)
        axs[plotNo//5, plotNo%5].tick_params(axis='x', labelsize=6)
        axs[plotNo//5, plotNo%5].tick_params(axis='y', labelsize=6)
    # legend plot:
    axs[0,4].plot(2016,0,color=MyColorCycle[4,:], lw=0.8, linestyle='-') # System GHG, sector, Baseline, full RES
    axs[0,4].plot(2016,0,color=MyColorCycle[7,:], lw=0.8, linestyle='-') # System GHG, sector, RCP2.6, full RES
    axs[0,4].plot(2016,0,color=MyColorCycle[4,:], lw=0.6, linestyle='--') # System GHG, sector, Baseline, no RES
    axs[0,4].plot(2016,0,color=MyColorCycle[7,:], lw=0.6, linestyle='--') # System GHG, sector, RCP2.6, no RES
    axs[0,4].legend(LegendLables,shadow = False, prop={'size':5}, loc = 'upper right')         
    axs[0,3].axis('off')
    axs[0,4].axis('off')
    fig.suptitle(fig_titles[Sect], fontsize=14)
    for xm in range(0,5):
        plt.sca(axs[4,xm])
        plt.xticks([2020,2030,2040,2050,2060], ['2020','2030','2040','2050','2060'], rotation =90, fontsize = 6, fontweight = 'normal')
    plt.show()
    fig_name = fig_names[Sect]
    fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = PlotExpResolution, bbox_inches='tight')  

#
#
#
#
#
#
#
#