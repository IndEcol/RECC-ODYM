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
mywb2 = openpyxl.load_workbook(os.path.join(RECC_Paths.results_path,'RECC_Global_Results_Template_CascSens.xlsx')) # for material-related emissions
mywb3 = openpyxl.load_workbook(os.path.join(RECC_Paths.results_path,'RECC_Global_Results_Template_CascSens.xlsx')) # for material-related emissions with recycling credit AND forest carbon uptake
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
        Vsheet2          = mywb2[RegionalScope + '_Vehicles']
        Vsheet3          = mywb3[RegionalScope + '_Vehicles']
        NE               = 7 # 7 for vehs. and 6 for buildings
        
    if Setting == 'Cascade_reb':        
        CascadeFlag1     = True
        SectorString     = 'reb'
        Vsheet           = mywb[RegionalScope  + '_ResBuildings']
        Vsheet2          = mywb2[RegionalScope + '_ResBuildings']
        Vsheet3          = mywb3[RegionalScope + '_ResBuildings']
        NE               = 6 # 7 for vehs. and 6 for buildings     

    if Setting == 'Cascade_nrb':        
        CascadeFlag1     = True
        SectorString     = 'nrb'
        Vsheet           = mywb[RegionalScope  + '_NonResBuildings']
        Vsheet2          = mywb2[RegionalScope + '_NonResBuildings']
        Vsheet3          = mywb3[RegionalScope + '_NonResBuildings']
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
        ASummary, AvgDecadalEms, MatSummary, AvgDecadalMatEms, MatSummaryC, AvgDecadalMatEmsC, CumEms, AnnEms2050, MatStocks, TimeSeries_R = ODYM_RECC_Cascade_V2_4.main(RegionalScope,FolderList,SectorString)
        # write results summary to Excel
        for R in range(0,NR):
            for r in range(0,3):
                for c in range(0,NE):
                    Vsheet.cell(row = r+3,  column = c +CascCols[R]).value   = ASummary[r,R,c]       
                    Vsheet.cell(row = r+9,  column = c +CascCols[R]).value   = ASummary[r+3,R,c]       
                    Vsheet.cell(row = r+15, column = c +CascCols[R]).value   = ASummary[r+6,R,c]
                    Vsheet.cell(row = r+38, column = c +CascCols[R]).value   = ASummary[r+9,R,c]
                    for d in range(0,4):
                        Vsheet.cell(row = d*3 + r + 21,column = c +CascCols[R]).value  = AvgDecadalEms[r,R,c,d]
            for r in range(0,3):
                Vsheet2.cell(row = r+3,  column = 4).value  = 0       
                Vsheet2.cell(row = r+9,  column = 4).value  = 0       
                for c in range(0,NE):
                    Vsheet2.cell(row = r+3,  column = c +CascCols[R]).value  = MatSummary[r,R,c]       
                    Vsheet2.cell(row = r+9,  column = c +CascCols[R]).value  = MatSummary[r+3,R,c]       
                    Vsheet2.cell(row = r+15, column = c +CascCols[R]).value  = MatSummary[r+6,R,c]
                    Vsheet2.cell(row = r+38, column = c +CascCols[R]).value  = MatSummary[r+9,R,c]
                    for d in range(0,4):
                        Vsheet2.cell(row = d*3 + r + 21,column = c +CascCols[R]).value  = AvgDecadalMatEms[r,R,c,d]                    
            for r in range(0,3):
                Vsheet3.cell(row = r+3,  column = 4).value  = 0       
                Vsheet3.cell(row = r+9,  column = 4).value  = 0       
                for c in range(0,NE):
                    Vsheet3.cell(row = r+3,  column = c +CascCols[R]).value  = MatSummaryC[r,R,c]       
                    Vsheet3.cell(row = r+9,  column = c +CascCols[R]).value  = MatSummaryC[r+3,R,c]       
                    Vsheet3.cell(row = r+15, column = c +CascCols[R]).value  = MatSummaryC[r+6,R,c]
                    Vsheet3.cell(row = r+38, column = c +CascCols[R]).value  = MatSummaryC[r+9,R,c]
                    for d in range(0,4):
                        Vsheet3.cell(row = d*3 + r + 21,column = c +CascCols[R]).value  = AvgDecadalMatEmsC[r,R,c,d]   
        
        # Store results in time series array
        if SectorString == 'pav' or SectorString == 'reb':
            if SectorString == 'pav':
                SectorIndex = 0
            if SectorString == 'reb':
                SectorIndex = 1
            RegPos = RegionList20.index(RegionalScope)
            PlotOrder.append(RegPos)
            TimeSeries_All[0,:,RegPos,SectorIndex,:,:] = TimeSeries_R[0,0,:,:,:]
            TimeSeries_All[1,:,RegPos,SectorIndex,:,:] = TimeSeries_R[0,-1,:,:,:]
            TimeSeries_All[2,:,RegPos,SectorIndex,:,:] = TimeSeries_R[1,0,:,:,:]
            TimeSeries_All[3,:,RegPos,SectorIndex,:,:] = TimeSeries_R[1,-1,:,:,:]
        
        RCP_Matstocks = 1 # MatStocks are plotted for RCP2.6 only
        
        if Setting == 'Cascade_pav':                    
            # store other results
            Table_Annual[0,0:-1,1,:]= AnnEms2050[1,:,:].transpose().copy()
            Table_CumEms[0,0:-1,1,:]= CumEms[1,:,:].transpose().copy()
            MatStocksTab1[0,:]    = MatStocks[4,:,0,RCP_Matstocks,0].copy()
            MatStocksTab1[1,:]    = MatStocks[34,:,0,RCP_Matstocks,0].copy()
            MatStocksTab1[2,:]    = MatStocks[34,:,0,RCP_Matstocks,-1].copy()
            MatStocksTab2[0,:]    = MatStocks[4,:,1,RCP_Matstocks,0].copy()
            MatStocksTab2[1,:]    = MatStocks[34,:,1,RCP_Matstocks,0].copy()
            MatStocksTab2[2,:]    = MatStocks[34,:,1,RCP_Matstocks,-1].copy()
            MatStocksTab3[0,:]    = MatStocks[4,:,2,RCP_Matstocks,0].copy()
            MatStocksTab3[1,:]    = MatStocks[34,:,2,RCP_Matstocks,0].copy()
            MatStocksTab3[2,:]    = MatStocks[34,:,2,RCP_Matstocks,-1].copy()
                    
        if Setting == 'Cascade_reb':
            # store other results
            Table_Annual[1,0:5,1,:] = AnnEms2050[1,:,0:-1].transpose().copy()
            Table_Annual[1,7,1,:]   = AnnEms2050[1,:,-1].copy()
            Table_CumEms[1,0:5,1,:] = CumEms[1,:,0:-1].transpose().copy()                    
            Table_CumEms[1,7,1,:]   = CumEms[1,:,-1].copy()        
            MatStocksTab1[3,:]    = MatStocks[4,:,0,RCP_Matstocks,0].copy()
            MatStocksTab1[4,:]    = MatStocks[34,:,0,RCP_Matstocks,0].copy()
            MatStocksTab1[5,:]    = MatStocks[34,:,0,RCP_Matstocks,-1].copy()            
            MatStocksTab2[3,:]    = MatStocks[4,:,1,RCP_Matstocks,0].copy()
            MatStocksTab2[4,:]    = MatStocks[34,:,1,RCP_Matstocks,0].copy()
            MatStocksTab2[5,:]    = MatStocks[34,:,1,RCP_Matstocks,-1].copy()            
            MatStocksTab3[3,:]    = MatStocks[4,:,2,RCP_Matstocks,0].copy()
            MatStocksTab3[4,:]    = MatStocks[34,:,2,RCP_Matstocks,0].copy()
            MatStocksTab3[5,:]    = MatStocks[34,:,2,RCP_Matstocks,-1].copy()            
    
        if Setting == 'Cascade_nrb':                        
            # store other results
            Table_Annual[1,0:5,2,:] = AnnEms2050[1,:,0:-1].transpose().copy()
            Table_Annual[1,7,2,:]   = AnnEms2050[1,:,-1].copy()
            Table_CumEms[1,0:5,2,:] = CumEms[1,:,0:-1].transpose().copy()                    
            Table_CumEms[1,7,2,:]   = CumEms[1,:,-1].copy()      
            MatStocksTab1[6,:]    = MatStocks[4,:,0,RCP_Matstocks,0].copy()
            MatStocksTab1[7,:]    = MatStocks[34,:,0,RCP_Matstocks,0].copy()
            MatStocksTab1[8,:]    = MatStocks[34,:,0,RCP_Matstocks,-1].copy()                 
            MatStocksTab2[6,:]    = MatStocks[4,:,1,RCP_Matstocks,0].copy()
            MatStocksTab2[7,:]    = MatStocks[34,:,1,RCP_Matstocks,0].copy()
            MatStocksTab2[8,:]    = MatStocks[34,:,1,RCP_Matstocks,-1].copy()                 
            MatStocksTab3[6,:]    = MatStocks[4,:,2,RCP_Matstocks,0].copy()
            MatStocksTab3[7,:]    = MatStocks[34,:,2,RCP_Matstocks,0].copy()
            MatStocksTab3[8,:]    = MatStocks[34,:,2,RCP_Matstocks,-1].copy()                 
                    
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
        ASummary, AvgDecadalEms, MatSummary, AvgDecadalMatEms, MatSummaryC, AvgDecadalMatEmsC, MatProduction_Prim, MatProduction_Sec = ODYM_RECC_Cascade_V2_4.main(RegionalScope,MultiSectorList,SectorString)
                    
            
    if Setting == 'Sensitivity_pav':
        SensitiFlag1     = True
        SectorString     = 'pav'
        NE               = 11 # 11 for vehs. and 10 for buildings 
        SensRows         = [4,9,14,19,24]
        
    if Setting == 'Sensitivity_reb':        
        SensitiFlag1     = True
        SectorString     = 'reb'
        NE               = 10 # 11 for vehs. and 10 for buildings        
        SensRows         = [40,45,50,55,60]
        
    if Setting == 'Sensitivity_nrb':        
        SensitiFlag1     = True
        SectorString     = 'nrb'
        NE               = 10 # 11 for vehs. and 10 for buildings        
        SensRows         = [76,81,86,91,96]
    
    SensCols             = [6,18] 
        
    if SensitiFlag1 is True:
        SensitiFlag1 = False
        for m in range(0,int(NE)):
            FolderList.append(ModelEvalListSheet.cell_value(Row +m, 3))
        # run the ODYM-RECC sensitivity analysis for pav
        CumEms_Sens, AnnEms2030_Sens, AnnEms2050_Sens, AvgDecadalEmsSens, MatCumEms_Sens, MatAnnEms2030_Sens, MatAnnEms2050_Sens, MatAvgDecadalEmsSens, MatCumEms_SensC, MatAnnEms2030_SensC, MatAnnEms2050_SensC, MatAvgDecadalEmsCSens, CumEms_Sens2060, MatCumEms_Sens2060, MatCumEms_SensC2060 = ODYM_RECC_Sensitivity_V2_4.main(RegionalScope,FolderList,SectorString)        
       
        # write results summary to Excel
        Ssheet  = mywb['Sensitivity_'  + RegionalScope]
        Ssheet2 = mywb2['Sensitivity_' + RegionalScope]
        Ssheet3 = mywb3['Sensitivity_' + RegionalScope]
        print('Sensitivity_' + RegionalScope + '_' + SectorString)
        for R in range(0,NR):
            for r in range(0,3):
                for c in range(0,NE):
                    Ssheet.cell(row = r +SensRows[0], column   = c +SensCols[R]).value   = AnnEms2030_Sens[r,R,c]
                    Ssheet.cell(row = r +SensRows[1], column   = c +SensCols[R]).value   = AnnEms2050_Sens[r,R,c]
                    Ssheet.cell(row = r +SensRows[2], column   = c +SensCols[R]).value   = CumEms_Sens[r,R,c]
                    Ssheet.cell(row = r +SensRows[3], column   = c +SensCols[R]).value   = CumEms_Sens2060[r,R,c]
                    for d in range(0,4):
                        Ssheet.cell(row = d*3 + r + SensRows[4],column  = c +SensCols[R]).value   = AvgDecadalEmsSens[r,R,c,d]
            
            for r in range(0,3):
                for c in range(0,NE):
                    Ssheet2.cell(row = r +SensRows[0], column  = c +SensCols[R]).value   = MatAnnEms2030_Sens[r,R,c]
                    Ssheet2.cell(row = r +SensRows[1], column  = c +SensCols[R]).value   = MatAnnEms2050_Sens[r,R,c]
                    Ssheet2.cell(row = r +SensRows[2], column  = c +SensCols[R]).value   = MatCumEms_Sens[r,R,c]
                    Ssheet2.cell(row = r +SensRows[3], column  = c +SensCols[R]).value   = MatCumEms_Sens2060[r,R,c]
                    for d in range(0,4):
                        Ssheet2.cell(row = d*3 + r + SensRows[4],column = c +SensCols[R]).value   = MatAvgDecadalEmsSens[r,R,c,d]       

            for r in range(0,3):
                for c in range(0,NE):
                    Ssheet3.cell(row = r +SensRows[0], column  = c +SensCols[R]).value   = MatAnnEms2030_SensC[r,R,c]
                    Ssheet3.cell(row = r +SensRows[1], column  = c +SensCols[R]).value   = MatAnnEms2050_SensC[r,R,c]
                    Ssheet3.cell(row = r +SensRows[2], column  = c +SensCols[R]).value   = MatCumEms_SensC[r,R,c]
                    Ssheet3.cell(row = r +SensRows[3], column  = c +SensCols[R]).value   = MatCumEms_SensC2060[r,R,c]
                    for d in range(0,4):
                        Ssheet3.cell(row = d*3 + r + SensRows[4],column = c +SensCols[R]).value   = MatAvgDecadalEmsCSens[r,R,c,d]                        

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
mywb2.save(os.path.join(RECC_Paths.results_path,'RECC_Global_Results_MaterialGHG_V2_4.xlsx'))    
mywb3.save(os.path.join(RECC_Paths.results_path,'RECC_Global_Results_MaterialGHG_inclRecyclingCredit_V2_4.xlsx'))        
mywb4.save(os.path.join(RECC_Paths.results_path,'RECC_Global_Results_Tables_V2_4.xlsx'))      
   
# plot SSP1 time series in 5x5 plot:
# TimeSeries_All indices: NX x Nt x Nr x NV x NS x NR / indicators x time x regions x sectors x SSP x RCP
# 0: system-wide GHG, no RES 1: system-wide GHG, all RES
# 2: material-related GHG, no RES, 3: material-related GHG, all RES,

# System-wide GHG
MyColorCycle = pylab.cm.Set1(np.arange(0,1,0.1)) # select 12 colors from the 'Paired' color map. 
plt.rcParams['axes.labelsize'] = 7
fig_names    = ['GHG_SSP1_pav_5x5.png','GHG_SSP1_reb_5x5.png']
fig_titles   = ['System-wide GHG, SSP1, pav','System-wide GHG, SSP1, reb']
LegendLables = ['RCP NoPol, full RES','RCP 2.6, full RES','RCP NoPol, no RES','RCP 2.6, no RES']
for Sect in range(0,2):
    fig, axs = plt.subplots(5, 5, sharex=True, gridspec_kw={'hspace': 0.6, 'wspace': 0.4})
    for plotNo in PlotOrder:
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[0,:,plotNo,Sect,1,0],color=MyColorCycle[4,:], lw=0.6, linestyle='--') # System GHG, sector, Baseline, no RES
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[0,:,plotNo,Sect,1,1],color=MyColorCycle[7,:], lw=0.6, linestyle='--') # System GHG, sector, RCP2.6, no RES
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[1,:,plotNo,Sect,1,0],color=MyColorCycle[4,:], lw=0.8, linestyle='-') # System GHG, sector, Baseline, full RES
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[1,:,plotNo,Sect,1,1],color=MyColorCycle[7,:], lw=0.8, linestyle='-') # System GHG, sector, RCP2.6, full RES
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
    fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = 400, bbox_inches='tight')  
    
 # Material cycle GHG
MyColorCycle = pylab.cm.Set1(np.arange(0,1,0.1)) # select 12 colors from the 'Paired' color map. 
plt.rcParams['axes.labelsize'] = 7
fig_names = ['GHGMat_SSP1_pav_5x5.png','GHGMat_SSP1_reb_5x5.png']
fig_titles = ['Matcycle GHG, SSP1, pav','Matcycle GHG, SSP1, reb']
LegendLables = ['RCP NoPol, full RES','RCP 2.6, full RES','RCP NoPol, no RES','RCP 2.6, no RES']
for Sect in range(0,2):
    fig, axs = plt.subplots(5, 5, sharex=True, gridspec_kw={'hspace': 0.6, 'wspace': 0.4})
    for plotNo in PlotOrder:
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[2,:,plotNo,Sect,1,0],color=MyColorCycle[4,:], lw=0.6, linestyle='--') # System GHG, sector, Baseline, no RES
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[2,:,plotNo,Sect,1,1],color=MyColorCycle[7,:], lw=0.6, linestyle='--') # System GHG, sector, RCP2.6, no RES
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[3,:,plotNo,Sect,1,0],color=MyColorCycle[4,:], lw=0.8, linestyle='-') # System GHG, sector, Baseline, full RES
        axs[plotNo//5, plotNo%5].plot(np.arange(2016,2061), TimeSeries_All[3,:,plotNo,Sect,1,1],color=MyColorCycle[7,:], lw=0.8, linestyle='-') # System GHG, sector, RCP2.6, full RES
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
    fig.savefig(os.path.join(RECC_Paths.results_path,fig_name), dpi = 400, bbox_inches='tight')    



#
#
#
#
#
#
#
#