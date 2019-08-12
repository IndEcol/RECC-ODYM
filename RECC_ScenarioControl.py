# -*- coding: utf-8 -*-
"""
Created on Fri Jun 14 05:18:48 2019

@author: spauliuk
"""

"""

File RECC_ScenarioControl.py

Script that modifies the RECC config file to run a list of scenarios and executes RECC main script for each scenario config.

"""

# Import required libraries:
import os
import xlrd
import openpyxl

import RECC_Paths # Import path file
import RECC_G7IC_V2_1

#ScenarioSetting, sheet name of RECC_ModelConfig_List.xlsx to be selected:
#ScenarioSetting = 'RECC_Config_Cascade'
ScenarioSetting = 'RECC_Config_Sensitivity'
#ScenarioSetting = 'SingleTestRun'
#ScenarioSetting = 'GroupTestRun'


# open scenario sheet
ModelConfigListFile  = xlrd.open_workbook(os.path.join(RECC_Paths.recc_path,'RECC_ModelConfig_List_V2_1.xlsx'))
ModelConfigListSheet = ModelConfigListFile.sheet_by_name(ScenarioSetting)

#Read control lines and execute main model script
ResultFolders = []
Row = 3
# search for script config list entry
while True:
    try:
        SheetName = ModelConfigListSheet.cell_value(Row, 2)
        print(SheetName)
        Config = {}
        for m in range(3,19):
            Config[ModelConfigListSheet.cell_value(2, m)] = ModelConfigListSheet.cell_value(Row, m)
    except:
        break
    Row += 1
    # rewrite RECC model config
    mywb = openpyxl.load_workbook(os.path.join(RECC_Paths.recc_path,'RECC_Config_V2_0.xlsx'))
    
    sheet = mywb.get_sheet_by_name('Config')
    sheet['D4'] = SheetName
    sheet = mywb.get_sheet_by_name(SheetName)
    sheet['D131'] = Config['Logging_Verbosity']
    sheet['D132'] = Config['Include_REStrategy_FabYieldImprovement']
    sheet['D133'] = Config['Include_REStrategy_FabScrapDiversion']
    sheet['D134'] = Config['Include_REStrategy_EoL_RR_Improvement']
    sheet['D135'] = Config['ScrapExport']
    sheet['D136'] = Config['ScrapExportRecyclingCredit']
    sheet['D137'] = Config['IncludeRecycling']
    sheet['D138'] = Config['Include_REStrategy_MaterialSubstitution']
    sheet['D139'] = Config['Include_REStrategy_UsingLessMaterialByDesign']
    sheet['D140'] = Config['Include_REStrategy_ReUse']
    sheet['D141'] = Config['Include_REStrategy_LifeTimeExtension']
    sheet['D142'] = Config['Include_REStrategy_MoreIntenseUse']
    sheet['D143'] = Config['Include_REStrategy_CarSharing']
    sheet['D144'] = Config['Include_REStrategy_RideSharing']
    sheet['D145'] = Config['Include_REStrategy_ModalSplit']
    sheet['D146'] = Config['SectorSelect']
    
    mywb.save(os.path.join(RECC_Paths.recc_path,'RECC_Config_V2_0.xlsx'))
    # run the ODYM-RECC model
    ResultFolder = RECC_G7IC_V2_1.main()
    ResultFolders.append(ResultFolder)




#
#