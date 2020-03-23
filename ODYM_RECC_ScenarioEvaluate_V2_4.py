# -*- coding: utf-8 -*-
"""
Created on Fri Jun 14 05:18:48 2019

@author: spauliuk
"""

"""

File RECC_ScenarioEvaluate_V2_3.py

Script that runs the sensitivity and scnenario comparison scripts for different settings.

"""

# Import required libraries:
import os
import xlrd
import openpyxl
import numpy as np

import RECC_Paths # Import path file
import ODYM_RECC_Cascade_PassVehicles_V2_3
import ODYM_RECC_Cascade_ResBuildings_V2_3
import ODYM_RECC_Sensitivity_PassVehicles_V2_3
import ODYM_RECC_Sensitivity_ResBuildings_V2_3
import ODYM_RECC_v2_3_Table_Extract
import ODYM_RECC_Cascade_PAV_REB_NRB_V2_3
import ODYM_RECC_Cascade_Efficiency_Sufficiency_V2_3

#ScenarioSetting, sheet name of RECC_ModelConfig_List.xlsx to be selected:
ScenarioSetting = 'Evaluate_RECC_Cascade'
#ScenarioSetting = 'Evaluate_GroupTestRun'
#ScenarioSetting = 'Evaluate_RECC_Sensitivity'

# open scenario sheet
ModelConfigListFile  = xlrd.open_workbook(os.path.join(RECC_Paths.recc_path,'RECC_ModelConfig_List_V2_3.xlsx'))
ModelEvalListSheet = ModelConfigListFile.sheet_by_name(ScenarioSetting)

# open result summary file
mywb  = openpyxl.load_workbook(os.path.join(RECC_Paths.results_path,'RECC_Germany_Results_Template.xlsx')) # for total emissions
mywb2 = openpyxl.load_workbook(os.path.join(RECC_Paths.results_path,'RECC_Germany_Results_Template.xlsx')) # for material-related emissions
mywb3 = openpyxl.load_workbook(os.path.join(RECC_Paths.results_path,'RECC_Germany_Results_Template.xlsx')) # for material-related emissions with recycling credit
mywb4 = openpyxl.load_workbook(os.path.join(RECC_Paths.results_path,'RECC_Germany_Results_Tables_Template.xlsx')) # for emissions to be reported in Table 2.

#Read control lines and execute main model script
Row  = 1

NoofCascadeSteps_pav     = 7 # 7 for vehs. and 6 for buildings
NoofCascadeSteps_reb     = 6 # 7 for vehs. and 6 for buildings
NoofCascadeSteps_nrb     = 6 # 7 for vehs. and 6 for buildings
NoofCascadeSteps_pnr     = 8 # 8 for vehs, res and nonres buildings
NoofSensitivitySteps_pav = 11 # no of different sensitivity analysis cases for pav
NoofSensitivitySteps_reb = 10 # no of different sensitivity analysis cases for reb
NoofSensitivitySteps_nrb = 10 # no of different sensitivity analysis cases for reb

Table2_Annual = np.zeros((8,3))
Table2_CumEms = np.zeros((8,3))
MatStocksTab1 = np.zeros((9,6))
MatStocksTab2 = np.zeros((9,6))
MatStocksTab3 = np.zeros((9,6))

SingleSectList = []
# search for script config list entry
while ModelEvalListSheet.cell_value(Row, 1) != 'ENDOFLIST':
    if ModelEvalListSheet.cell_value(Row, 1) != '':
        PassVehList    = []
        ResBldsList    = []
        NonResBldsList = []
        ThreeSectoList = []
        RegionalScope  = ModelEvalListSheet.cell_value(Row, 1)
        Setting        = ModelEvalListSheet.cell_value(Row, 2) # cascade or sensitivity
        print(RegionalScope)
        
    if Setting == 'Cascade_pav':
        for m in range(0,int(NoofCascadeSteps_pav)):
            PassVehList.append(ModelEvalListSheet.cell_value(Row +m, 3))
        # run the ODYM-RECC scenario comparison
        ASummaryV, AvgDecadalEmsV, MatSummaryV, AvgDecadalMatEmsV, MatSummaryVC, AvgDecadalMatEmsVC, CumEmsV, AnnEmsV2050, MatStocks = ODYM_RECC_Cascade_PassVehicles_V2_3.main(RegionalScope,PassVehList)
        # write results summary to Excel
        Vsheet = mywb[RegionalScope + '_Vehicles']
        for r in range(0,3):
            for c in range(0,7):
                Vsheet.cell(row = r+3,  column = c+5).value  = ASummaryV[r,c]       
                Vsheet.cell(row = r+9,  column = c+5).value  = ASummaryV[r+3,c]       
                Vsheet.cell(row = r+15, column = c+5).value  = ASummaryV[r+6,c]
                Vsheet.cell(row = r+38, column = c+5).value  = ASummaryV[r+9,c]
                for d in range(0,4):
                    Vsheet.cell(row = d*3 + r + 21,column = c+5).value  = AvgDecadalEmsV[r,c,d]
        Vsheet2 = mywb2[RegionalScope + '_Vehicles']
        for r in range(0,3):
            Vsheet2.cell(row = r+3,  column = 4).value  = 0       
            Vsheet2.cell(row = r+9,  column = 4).value  = 0       
            for c in range(0,7):
                Vsheet2.cell(row = r+3,  column = c+5).value  = MatSummaryV[r,c]       
                Vsheet2.cell(row = r+9,  column = c+5).value  = MatSummaryV[r+3,c]       
                Vsheet2.cell(row = r+15, column = c+5).value  = MatSummaryV[r+6,c]
                Vsheet2.cell(row = r+38, column = c+5).value  = MatSummaryV[r+9,c]
                for d in range(0,4):
                    Vsheet2.cell(row = d*3 + r + 21,column = c+5).value  = AvgDecadalMatEmsV[r,c,d]                    
        Vsheet3 = mywb3[RegionalScope + '_Vehicles']
        for r in range(0,3):
            Vsheet3.cell(row = r+3,  column = 4).value  = 0       
            Vsheet3.cell(row = r+9,  column = 4).value  = 0       
            for c in range(0,7):
                Vsheet3.cell(row = r+3,  column = c+5).value  = MatSummaryVC[r,c]       
                Vsheet3.cell(row = r+9,  column = c+5).value  = MatSummaryVC[r+3,c]       
                Vsheet3.cell(row = r+15, column = c+5).value  = MatSummaryVC[r+6,c]
                Vsheet3.cell(row = r+38, column = c+5).value  = MatSummaryVC[r+9,c]
                for d in range(0,4):
                    Vsheet3.cell(row = d*3 + r + 21,column = c+5).value  = AvgDecadalMatEmsVC[r,c,d]   
                    
        # store other results
        Table2_Annual[0:-1,0] = AnnEmsV2050[1,1,:].copy()
        Table2_CumEms[0:-1,0] = CumEmsV[1,1,:].copy()
        MatStocksTab1[0,:]    = MatStocks[4,:,0,1,0].copy()
        MatStocksTab1[1,:]    = MatStocks[34,:,0,1,0].copy()
        MatStocksTab1[2,:]    = MatStocks[34,:,0,1,-1].copy()
        MatStocksTab2[0,:]    = MatStocks[4,:,1,1,0].copy()
        MatStocksTab2[1,:]    = MatStocks[34,:,1,1,0].copy()
        MatStocksTab2[2,:]    = MatStocks[34,:,1,1,-1].copy()
        MatStocksTab3[0,:]    = MatStocks[4,:,2,1,0].copy()
        MatStocksTab3[1,:]    = MatStocks[34,:,2,1,0].copy()
        MatStocksTab3[2,:]    = MatStocks[34,:,2,1,-1].copy()
                
    if Setting == 'Cascade_reb':
        for m in range(0,int(NoofCascadeSteps_reb)):
            ResBldsList.append(ModelEvalListSheet.cell_value(Row +m, 3))
        # run the ODYM-RECC scenario comparison
        ASummaryB, AvgDecadalEmsB, MatSummaryB, AvgDecadalMatEmsB, MatSummaryBC, AvgDecadalMatEmsBC, CumEmsV, AnnEmsV2050, MatStocks = ODYM_RECC_Cascade_ResBuildings_V2_3.main(RegionalScope,ResBldsList,['Residential_buildings'])
        # write results summary to Excel
        Bsheet = mywb[RegionalScope + '_ResBuildings']
        for r in range(0,3):
            for c in range(0,6):
                Bsheet.cell(row = r+3, column = c+5).value  = ASummaryB[r,c]       
                Bsheet.cell(row = r+9, column = c+5).value  = ASummaryB[r+3,c]       
                Bsheet.cell(row = r+15, column = c+5).value = ASummaryB[r+6,c]     
                Bsheet.cell(row = r+38, column = c+5).value = ASummaryB[r+9,c]     
                for d in range(0,4):
                    Bsheet.cell(row = d*3 + r + 21,column = c+5).value  = AvgDecadalEmsB[r,c,d]
        Bsheet2 = mywb2[RegionalScope + '_ResBuildings']
        for r in range(0,3):
            Bsheet2.cell(row = r+3,  column = 4).value  = 0       
            Bsheet2.cell(row = r+9,  column = 4).value  = 0                   
            for c in range(0,6):
                Bsheet2.cell(row = r+3,  column = c+5).value  = MatSummaryB[r,c]       
                Bsheet2.cell(row = r+9,  column = c+5).value  = MatSummaryB[r+3,c]       
                Bsheet2.cell(row = r+15, column = c+5).value  = MatSummaryB[r+6,c]
                Bsheet2.cell(row = r+38, column = c+5).value  = MatSummaryB[r+9,c]
                for d in range(0,4):
                    Bsheet2.cell(row = d*3 + r + 21,column = c+5).value  = AvgDecadalMatEmsB[r,c,d]    
        Bsheet3 = mywb3[RegionalScope + '_ResBuildings']
        for r in range(0,3):
            Bsheet3.cell(row = r+3,  column = 4).value  = 0       
            Bsheet3.cell(row = r+9,  column = 4).value  = 0                   
            for c in range(0,6):
                Bsheet3.cell(row = r+3,  column = c+5).value  = MatSummaryBC[r,c]       
                Bsheet3.cell(row = r+9,  column = c+5).value  = MatSummaryBC[r+3,c]       
                Bsheet3.cell(row = r+15, column = c+5).value  = MatSummaryBC[r+6,c]
                Bsheet3.cell(row = r+38, column = c+5).value  = MatSummaryBC[r+9,c]
                for d in range(0,4):
                    Bsheet3.cell(row = d*3 + r + 21,column = c+5).value  = AvgDecadalMatEmsBC[r,c,d]     

        # store other results
        Table2_Annual[0:5,1] = AnnEmsV2050[1,1,0:-1].copy()
        Table2_Annual[7,1]   = AnnEmsV2050[1,1,-1].copy()
        Table2_CumEms[0:5,1] = CumEmsV[1,1,0:-1].copy()                    
        Table2_CumEms[7,1]   = CumEmsV[1,1,-1].copy()        
        MatStocksTab1[3,:]    = MatStocks[4,:,0,1,0].copy()
        MatStocksTab1[4,:]    = MatStocks[34,:,0,1,0].copy()
        MatStocksTab1[5,:]    = MatStocks[34,:,0,1,-1].copy()            
        MatStocksTab2[3,:]    = MatStocks[4,:,1,1,0].copy()
        MatStocksTab2[4,:]    = MatStocks[34,:,1,1,0].copy()
        MatStocksTab2[5,:]    = MatStocks[34,:,1,1,-1].copy()            
        MatStocksTab3[3,:]    = MatStocks[4,:,2,1,0].copy()
        MatStocksTab3[4,:]    = MatStocks[34,:,2,1,0].copy()
        MatStocksTab3[5,:]    = MatStocks[34,:,2,1,-1].copy()            

    if Setting == 'Cascade_nrb':
        for m in range(0,int(NoofCascadeSteps_reb)):
            NonResBldsList.append(ModelEvalListSheet.cell_value(Row +m, 3))
        # run the ODYM-RECC scenario comparison
        ASummaryN, AvgDecadalEmsN, MatSummaryN, AvgDecadalMatEmsN, MatSummaryNC, AvgDecadalMatEmsNC, CumEmsV, AnnEmsV2050, MatStocks = ODYM_RECC_Cascade_ResBuildings_V2_3.main(RegionalScope,NonResBldsList,['Non-residential_buildings'])
        # write results summary to Excel
        Bsheet = mywb[RegionalScope + '_NonResBuildings']
        for r in range(0,3):
            for c in range(0,6):
                Bsheet.cell(row = r+3, column = c+5).value  = ASummaryN[r,c]       
                Bsheet.cell(row = r+9, column = c+5).value  = ASummaryN[r+3,c]       
                Bsheet.cell(row = r+15, column = c+5).value = ASummaryN[r+6,c]     
                Bsheet.cell(row = r+38, column = c+5).value = ASummaryN[r+9,c]     
                for d in range(0,4):
                    Bsheet.cell(row = d*3 + r + 21,column = c+5).value  = AvgDecadalEmsN[r,c,d]
        Bsheet2 = mywb2[RegionalScope + '_NonResBuildings']
        for r in range(0,3):
            Bsheet2.cell(row = r+3,  column = 4).value  = 0       
            Bsheet2.cell(row = r+9,  column = 4).value  = 0                   
            for c in range(0,6):
                Bsheet2.cell(row = r+3,  column = c+5).value  = MatSummaryN[r,c]       
                Bsheet2.cell(row = r+9,  column = c+5).value  = MatSummaryN[r+3,c]       
                Bsheet2.cell(row = r+15, column = c+5).value  = MatSummaryN[r+6,c]
                Bsheet2.cell(row = r+38, column = c+5).value  = MatSummaryN[r+9,c]
                for d in range(0,4):
                    Bsheet2.cell(row = d*3 + r + 21,column = c+5).value  = AvgDecadalMatEmsN[r,c,d]    
        Bsheet3 = mywb3[RegionalScope + '_NonResBuildings']
        for r in range(0,3):
            Bsheet3.cell(row = r+3,  column = 4).value  = 0       
            Bsheet3.cell(row = r+9,  column = 4).value  = 0                   
            for c in range(0,6):
                Bsheet3.cell(row = r+3,  column = c+5).value  = MatSummaryNC[r,c]       
                Bsheet3.cell(row = r+9,  column = c+5).value  = MatSummaryNC[r+3,c]       
                Bsheet3.cell(row = r+15, column = c+5).value  = MatSummaryNC[r+6,c]
                Bsheet3.cell(row = r+38, column = c+5).value  = MatSummaryNC[r+9,c]
                for d in range(0,4):
                    Bsheet3.cell(row = d*3 + r + 21,column = c+5).value  = AvgDecadalMatEmsNC[r,c,d]  
                    
        # store other results
        Table2_Annual[0:5,2] = AnnEmsV2050[1,1,0:-1].copy()
        Table2_Annual[7,2]   = AnnEmsV2050[1,1,-1].copy()
        Table2_CumEms[0:5,2] = CumEmsV[1,1,0:-1].copy()                    
        Table2_CumEms[7,2]   = CumEmsV[1,1,-1].copy()      
        MatStocksTab1[6,:]    = MatStocks[4,:,0,1,0].copy()
        MatStocksTab1[7,:]    = MatStocks[34,:,0,1,0].copy()
        MatStocksTab1[8,:]    = MatStocks[34,:,0,1,-1].copy()                 
        MatStocksTab2[6,:]    = MatStocks[4,:,1,1,0].copy()
        MatStocksTab2[7,:]    = MatStocks[34,:,1,1,0].copy()
        MatStocksTab2[8,:]    = MatStocks[34,:,1,1,-1].copy()                 
        MatStocksTab3[6,:]    = MatStocks[4,:,2,1,0].copy()
        MatStocksTab3[7,:]    = MatStocks[34,:,2,1,0].copy()
        MatStocksTab3[8,:]    = MatStocks[34,:,2,1,-1].copy()                 
                    
    if Setting == 'Sensitivity_pav':
        for m in range(0,int(NoofSensitivitySteps_pav)):
            PassVehList.append(ModelEvalListSheet.cell_value(Row +m, 3))
        # run the ODYM-RECC sensitivity analysis for pav
        CumEmsV_Sens, AnnEmsV2030_Sens, AnnEmsV2050_Sens, AvgDecadalEms, MatCumEmsV_Sens, MatAnnEmsV2030_Sens, MatAnnEmsV2050_Sens, MatAvgDecadalEmsV, MatCumEmsV_SensC, MatAnnEmsV2030_SensC, MatAnnEmsV2050_SensC, MatAvgDecadalEmsC, CumEmsV_Sens2060, MatCumEmsV_Sens2060, MatCumEmsV_SensC2060 = ODYM_RECC_Sensitivity_PassVehicles_V2_3.main(RegionalScope,PassVehList)        
        # write results summary to Excel
        Ssheet = mywb['Sensitivity_' + RegionalScope]
        print(RegionalScope)
        for r in range(0,3):
            for c in range(0,11):
                Ssheet.cell(row = r+4, column = c+6).value  = AnnEmsV2030_Sens[r,c]
                Ssheet.cell(row = r+9, column = c+6).value  = AnnEmsV2050_Sens[r,c]
                Ssheet.cell(row = r+14,column = c+6).value  = CumEmsV_Sens[r,c]
                Ssheet.cell(row = r+19,column = c+6).value  = CumEmsV_Sens2060[r,c]
                for d in range(0,4):
                    Ssheet.cell(row = d*3 + r + 24,column = c+6).value  = AvgDecadalEms[r,c,d]
        Ssheet2 = mywb2['Sensitivity_' + RegionalScope]
        print(RegionalScope)
        for r in range(0,3):
            for c in range(0,11):
                Ssheet2.cell(row = r+4, column = c+6).value  = MatAnnEmsV2030_Sens[r,c]
                Ssheet2.cell(row = r+9, column = c+6).value  = MatAnnEmsV2050_Sens[r,c]
                Ssheet2.cell(row = r+14,column = c+6).value  = MatCumEmsV_Sens[r,c]
                Ssheet2.cell(row = r+19,column = c+6).value  = MatCumEmsV_Sens2060[r,c]
                for d in range(0,4):
                    Ssheet2.cell(row = d*3 + r + 24,column = c+6).value  = MatAvgDecadalEmsV[r,c,d]       
        Ssheet3 = mywb3['Sensitivity_' + RegionalScope]
        print(RegionalScope)
        for r in range(0,3):
            for c in range(0,11):
                Ssheet3.cell(row = r+4, column = c+6).value  = MatAnnEmsV2030_SensC[r,c]
                Ssheet3.cell(row = r+9, column = c+6).value  = MatAnnEmsV2050_SensC[r,c]
                Ssheet3.cell(row = r+14,column = c+6).value  = MatCumEmsV_SensC[r,c]
                Ssheet3.cell(row = r+19,column = c+6).value  = MatCumEmsV_SensC2060[r,c]
                for d in range(0,4):
                    Ssheet3.cell(row = d*3 + r + 24,column = c+6).value  = MatAvgDecadalEmsC[r,c,d]                        
                    
    if Setting == 'Sensitivity_reb':
        for m in range(0,int(NoofSensitivitySteps_reb)):
            ResBldsList.append(ModelEvalListSheet.cell_value(Row +m, 3))
        # run the ODYM-RECC sensitivity analysis for reb
        CumEmsB_Sens, AnnEmsB2030_Sens, AnnEmsB2050_Sens, AvgDecadalEms, MatCumEmsB_Sens, MatAnnEmsB2030_Sens, MatAnnEmsB2050_Sens, MatAvgDecadalEmsB, MatCumEmsV_SensC, MatAnnEmsV2030_SensC, MatAnnEmsV2050_SensC, MatAvgDecadalEmsC, CumEmsB_Sens2060, MatCumEmsV_Sens2060, MatCumEmsV_SensC2060 = ODYM_RECC_Sensitivity_ResBuildings_V2_3.main(RegionalScope,ResBldsList,['Residential buildings'])        
        # write results summary to Excel
        Ssheet = mywb['Sensitivity_' + RegionalScope]
        print(RegionalScope)
        for r in range(0,3):
            for c in range(0,10):
                Ssheet.cell(row = r+40,column = c+6).value  = AnnEmsB2030_Sens[r,c]
                Ssheet.cell(row = r+45,column = c+6).value  = AnnEmsB2050_Sens[r,c]
                Ssheet.cell(row = r+50,column = c+6).value  = CumEmsB_Sens[r,c]        
                Ssheet.cell(row = r+55,column = c+6).value  = CumEmsB_Sens2060[r,c]        
                for d in range(0,4):
                    Ssheet.cell(row = d*3 + r + 60,column = c+6).value  = AvgDecadalEms[r,c,d]   
        Ssheet2 = mywb2['Sensitivity_' + RegionalScope]
        print(RegionalScope)
        for r in range(0,3):
            for c in range(0,10):
                Ssheet2.cell(row = r+40, column = c+6).value  = MatAnnEmsB2030_Sens[r,c]
                Ssheet2.cell(row = r+45, column = c+6).value  = MatAnnEmsB2050_Sens[r,c]
                Ssheet2.cell(row = r+50,column = c+6).value   = MatCumEmsB_Sens[r,c]
                Ssheet2.cell(row = r+55,column = c+6).value   = MatCumEmsV_Sens2060[r,c]        
                for d in range(0,4):
                    Ssheet2.cell(row = d*3 + r + 60,column = c+6).value  = MatAvgDecadalEmsB[r,c,d]                      
        Ssheet3 = mywb3['Sensitivity_' + RegionalScope]
        print(RegionalScope)
        for r in range(0,3):
            for c in range(0,10):
                Ssheet3.cell(row = r+40, column = c+6).value  = MatAnnEmsV2030_SensC[r,c]
                Ssheet3.cell(row = r+45, column = c+6).value  = MatAnnEmsV2050_SensC[r,c]
                Ssheet3.cell(row = r+50,column = c+6).value   = MatCumEmsV_SensC[r,c]
                Ssheet3.cell(row = r+55,column = c+6).value   = MatCumEmsV_SensC2060[r,c]        
                for d in range(0,4):
                    Ssheet3.cell(row = d*3 + r + 60,column = c+6).value  = MatAvgDecadalEmsC[r,c,d] 
                    
    if Setting == 'Sensitivity_nrb':
        for m in range(0,int(NoofSensitivitySteps_reb)):
            ResBldsList.append(ModelEvalListSheet.cell_value(Row +m, 3))
        # run the ODYM-RECC sensitivity analysis for reb
        CumEmsB_Sens, AnnEmsB2030_Sens, AnnEmsB2050_Sens, AvgDecadalEms, MatCumEmsB_Sens, MatAnnEmsB2030_Sens, MatAnnEmsB2050_Sens, MatAvgDecadalEmsB, MatCumEmsV_SensC, MatAnnEmsV2030_SensC, MatAnnEmsV2050_SensC, MatAvgDecadalEmsC, CumEmsB_Sens2060, MatCumEmsV_Sens2060, MatCumEmsV_SensC2060 = ODYM_RECC_Sensitivity_ResBuildings_V2_3.main(RegionalScope,ResBldsList,['non-residential buildings'])        
        # write results summary to Excel
        Ssheet = mywb['Sensitivity_' + RegionalScope]
        print(RegionalScope)
        for r in range(0,3):
            for c in range(0,10):
                Ssheet.cell(row = r+76,column = c+6).value  = AnnEmsB2030_Sens[r,c]
                Ssheet.cell(row = r+81,column = c+6).value  = AnnEmsB2050_Sens[r,c]
                Ssheet.cell(row = r+86,column = c+6).value  = CumEmsB_Sens[r,c]        
                Ssheet.cell(row = r+91,column = c+6).value  = CumEmsB_Sens2060[r,c]        
                for d in range(0,4):
                    Ssheet.cell(row = d*3 + r + 96,column = c+6).value  = AvgDecadalEms[r,c,d]   
        Ssheet2 = mywb2['Sensitivity_' + RegionalScope]
        print(RegionalScope)
        for r in range(0,3):
            for c in range(0,10):
                Ssheet2.cell(row = r+76, column = c+6).value  = MatAnnEmsB2030_Sens[r,c]
                Ssheet2.cell(row = r+81, column = c+6).value  = MatAnnEmsB2050_Sens[r,c]
                Ssheet2.cell(row = r+86,column = c+6).value   = MatCumEmsB_Sens[r,c]
                Ssheet2.cell(row = r+91,column = c+6).value   = MatCumEmsV_Sens2060[r,c]        
                for d in range(0,4):
                    Ssheet2.cell(row = d*3 + r + 96,column = c+6).value  = MatAvgDecadalEmsB[r,c,d]                      
        Ssheet3 = mywb3['Sensitivity_' + RegionalScope]
        print(RegionalScope)
        for r in range(0,3):
            for c in range(0,10):
                Ssheet3.cell(row = r+76, column = c+6).value  = MatAnnEmsV2030_SensC[r,c]
                Ssheet3.cell(row = r+81, column = c+6).value  = MatAnnEmsV2050_SensC[r,c]
                Ssheet3.cell(row = r+86,column = c+6).value   = MatCumEmsV_SensC[r,c]
                Ssheet3.cell(row = r+91,column = c+6).value   = MatCumEmsV_SensC2060[r,c]        
                for d in range(0,4):
                    Ssheet3.cell(row = d*3 + r + 96,column = c+6).value  = MatAvgDecadalEmsC[r,c,d]                     
                    
    if Setting == 'Cascade_pav_reb_nrb':
        for m in range(0,int(NoofCascadeSteps_pnr)):
            ThreeSectoList.append(ModelEvalListSheet.cell_value(Row +m, 3))
        ThreeSectoList_Export = ThreeSectoList
        # run the ODYM-RECC scenario comparison  
        GHG_TableX = ODYM_RECC_v2_3_Table_Extract.main(RegionalScope,ThreeSectoList)        
        # write results summary as Table 2 to Excel
        Gsheet = mywb4['Table_X']
        print(RegionalScope)
        for r in range(0,4):
            for c in range(0,6):
                Gsheet.cell(row = r+4, column = c+4).value  = GHG_TableX[r,c]        

        # run the cascade plots for the three sectors
        ASummaryV, AvgDecadalEmsV, MatSummaryV, AvgDecadalMatEmsV, MatSummaryVC, AvgDecadalMatEmsVC, MatProduction_Prim, MatProduction_Sec = ODYM_RECC_Cascade_PAV_REB_NRB_V2_3.main(RegionalScope,ThreeSectoList)

    if Setting == 'Do_not_include':
        SingleSectList.append(ModelEvalListSheet.cell_value(Row, 3))

    # forward counter   
    if Setting == 'Cascade_pav':
        Row += NoofCascadeSteps_pav
    if Setting == 'Cascade_reb':
        Row += NoofCascadeSteps_reb        
    if Setting == 'Cascade_nrb':
        Row += NoofCascadeSteps_nrb        
    if Setting == 'Cascade_pav_reb_nrb':
        Row += NoofCascadeSteps_pnr
    if Setting == 'Sensitivity_pav':
        Row += NoofSensitivitySteps_pav
    if Setting == 'Sensitivity_reb':
        Row += NoofSensitivitySteps_reb
    if Setting == 'Sensitivity_nrb':
        Row += NoofSensitivitySteps_nrb    
    if Setting == 'Do_not_include':
        Row += 1    
        
                         
# run the efficieny_sufficieny plots (Fig. 6)
ODYM_RECC_Cascade_Efficiency_Sufficiency_V2_3.main(RegionalScope,ThreeSectoList_Export,SingleSectList)      

# store table 2:
WFsheet = mywb4['Table_2']
for u in range(0,8):
    for v in range(0,3):
        WFsheet.cell(row = u+4, column = v+3).value  = Table2_Annual[u,v]     
        WFsheet.cell(row = u+14, column = v+3).value  = Table2_CumEms[u,v]       
        
WFsheet = mywb4['Table_3']
for u in range(0,9):
    for v in range(0,6):
        WFsheet.cell(row = u+3 , column = v+6).value  = MatStocksTab1[u,v]     
        WFsheet.cell(row = u+14, column = v+6).value  = MatStocksTab2[u,v]     
        WFsheet.cell(row = u+25, column = v+6).value  = MatStocksTab3[u,v]     
    
mywb.save(os.path.join(RECC_Paths.results_path,'RECC_Germany_Results_SystemGHG_06_January_2020.xlsx'))
mywb2.save(os.path.join(RECC_Paths.results_path,'RECC_Germany_Results_MaterialGHG_06_January_2020.xlsx'))    
mywb3.save(os.path.join(RECC_Paths.results_path,'RECC_Germany_Results_MaterialGHG_inclRecyclingCredit_06_January_2020.xlsx'))        
mywb4.save(os.path.join(RECC_Paths.results_path,'RECC_Germany_Results_Tables.xlsx'))      
    
#
#
#
#
#
#
#
#