# -*- coding: utf-8 -*-
"""
Created on Jan 5th, 2020, as copy of RECC_ScenarioControl_V2_2.py

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
import ODYM_RECC_V2_4

#ScenarioSetting, sheet name of RECC_ModelConfig_List.xlsx to be selected:
ScenarioSetting = 'pav_reb_Config_list'
#ScenarioSetting = 'pav_reb_Config_list_all'
#ScenarioSetting = 'Germany_detail_config'
#ScenarioSetting = 'Germany_detail_config_all'
#ScenarioSetting = 'Global_all'
#ScenarioSetting = 'TestRun'

# open scenario sheet
ModelConfigListFile  = xlrd.open_workbook(os.path.join(RECC_Paths.recc_path,'RECC_ModelConfig_List_V2_4.xlsx'))
ModelConfigListSheet = ModelConfigListFile.sheet_by_name(ScenarioSetting)
SheetName = 'Config_Auto'
#Read control lines and execute main model script
ResultFolders = []
Row = 3
# search for script config list entry
while True:
    try:
        RegionalScope = ModelConfigListSheet.cell_value(Row, 2)
        print(RegionalScope)
        Config = {}
        for m in range(3,27):
            Config[ModelConfigListSheet.cell_value(2, m)] = ModelConfigListSheet.cell_value(Row, m)
    except:
        break
    Row += 1
    # rewrite RECC model config
    mywb = openpyxl.load_workbook(os.path.join(RECC_Paths.recc_path,'RECC_Config_V2_4.xlsx'))
    
    sheet = mywb.get_sheet_by_name('Cover')
    sheet['D4'] = SheetName
    sheet = mywb.get_sheet_by_name(SheetName)
    sheet['D7']   = RegionalScope
    sheet['G21']  = Config['RegionSelect']
    sheet['G27']  = Config['Products']
    sheet['G28']  = Config['Sectors']
    sheet['G29']  = Config['Products']
    sheet['G33']  = Config['NonresidentialBuildings']
    sheet['G48']  = Config['Regions32goods']
    sheet['D181'] = Config['Logging_Verbosity']
    sheet['D182'] = Config['Include_REStrategy_FabYieldImprovement']
    sheet['D183'] = Config['Include_REStrategy_FabScrapDiversion']
    sheet['D184'] = Config['Include_REStrategy_EoL_RR_Improvement']
    sheet['D185'] = Config['ScrapExport']
    sheet['D186'] = Config['ScrapExportRecyclingCredit']
    sheet['D187'] = Config['IncludeRecycling']
    sheet['D188'] = Config['Include_REStrategy_MaterialSubstitution']
    sheet['D189'] = Config['Include_REStrategy_UsingLessMaterialByDesign']
    sheet['D190'] = Config['Include_REStrategy_ReUse']
    sheet['D191'] = Config['Include_REStrategy_LifeTimeExtension']
    sheet['D192'] = Config['Include_REStrategy_MoreIntenseUse']
    sheet['D193'] = Config['Include_REStrategy_CarSharing']
    sheet['D194'] = Config['Include_REStrategy_RideSharing']
    sheet['D195'] = Config['Include_REStrategy_ModalSplit']
    sheet['D196'] = Config['SectorSelect']
    sheet['D197'] = Config['Include_Renovation_reb']
    sheet['D198'] = Config['Include_Renovation_nrb']
    sheet['D199'] = Config['No_EE_Improvements']
    
    mywb.save(os.path.join(RECC_Paths.recc_path,'RECC_Config_V2_4.xlsx'))
    # run the ODYM-RECC model
    OutputDict = ODYM_RECC_V2_4.main()
    ResultFolders.append(OutputDict['Name_Scenario'])


#
#