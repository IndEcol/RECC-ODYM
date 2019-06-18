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
import RECC_G7IC_V1_1

#ScenarioSetting, sheet name of RECC_ModelConfig_List.xlsx to be selected:
ScenarioSetting = 'RECC_Config_IRP_V1'
#ScenarioSetting = 'SingleTestRun'
#ScenarioSetting = 'GroupTestRun'


# open scenario sheet
ModelConfigListFile  = xlrd.open_workbook(os.path.join(RECC_Paths.recc_path,'RECC_ModelConfig_List.xlsx'))
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
        for m in range(3,14):
            Config[ModelConfigListSheet.cell_value(2, m)] = ModelConfigListSheet.cell_value(Row, m)
    except:
        break
    Row += 1
    # rewrite RECC model config
    mywb = openpyxl.load_workbook(os.path.join(RECC_Paths.recc_path,'RECC_Config.xlsx'))
    
    sheet = mywb.get_sheet_by_name('Config')
    sheet['D4'] = SheetName
    sheet = mywb.get_sheet_by_name(SheetName)
    sheet['D109'] = Config['Include_REStrategy_FabYieldImprovement']
    sheet['D110'] = Config['Include_REStrategy_EoL_RR_Improvement']
    sheet['D111'] = Config['Include_REStrategy_MaterialSubstitution']
    sheet['D112'] = Config['Include_REStrategy_UsingLessMaterialByDesign']
    sheet['D113'] = Config['Include_REStrategy_ReUse']
    sheet['D114'] = Config['Include_REStrategy_LifeTimeExtension']
    sheet['D115'] = Config['Include_REStrategy_MoreIntenseUse']
    sheet['D116'] = Config['Include_REStrategy_Sufficiency']
    sheet['D117'] = Config['SectorSelect']
    sheet['D118'] = Config['ScrapExportRecyclingCredit']
    sheet['D119'] = Config['IncludeRecycling']
    
    mywb.save(os.path.join(RECC_Paths.recc_path,'RECC_Config.xlsx'))
    # run the ODYM-RECC model
    ResultFolder = RECC_G7IC_V1_1.main()
    ResultFolders.append(ResultFolder)




#
#