# -*- coding: utf-8 -*-
"""
Created on Wed Oct 17 10:37:00 2018

@author: spauliuk
"""
def main(RegionalScope,ThreeSectoList):
    
    import xlrd
    import numpy as np
    import matplotlib.pyplot as plt  
    import os
    import RECC_Paths # Import path file   
    

    
    GHG_Table_Overview   = np.zeros((4,6,2)) # Table GHG overview format 4 scopes, 6 times, 2 RCP

    # No ME scenario
    Path = os.path.join(RECC_Paths.results_path,ThreeSectoList[0],'SysVar_TotalGHGFootprint.xls')
    Resultfile   = xlrd.open_workbook(Path)
    Resultsheet1 = Resultfile.sheet_by_name('Cover')
    UUID         = Resultsheet1.cell_value(3,2)
    Resultfile2  = xlrd.open_workbook(os.path.join(RECC_Paths.results_path,ThreeSectoList[0],'ODYM_RECC_ModelResults_' + UUID + '.xlsx'))
    Resultsheet2 = Resultfile2.sheet_by_name('Model_Results')
    # Find the index for the recycling credit and others:
    tci = 1
    while True:
        if Resultsheet2.cell_value(tci, 0) == 'GHG emissions, system-wide _3579di':
            break # that gives us the right index to read the recycling credit from the result table.
        tci += 1
    
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
        
    for nnn in range(0,2):
        GHG_Table_Overview[0,0,nnn] = Resultsheet2.cell_value(tci +2 + nnn,8)
        GHG_Table_Overview[0,1,nnn] = Resultsheet2.cell_value(tci +2 + nnn,12)
        GHG_Table_Overview[0,2,nnn] = Resultsheet2.cell_value(tci +2 + nnn,22)
        GHG_Table_Overview[0,3,nnn] = Resultsheet2.cell_value(tci +2 + nnn,32)
        GHG_Table_Overview[0,4,nnn] = Resultsheet2.cell_value(tci +2 + nnn,42)
        GHG_Table_Overview[0,5,nnn] = Resultsheet2.cell_value(tci +2 + nnn,52)
        
        GHG_Table_Overview[2,0,nnn] = Resultsheet2.cell_value(mci +2 + nnn,8)
        GHG_Table_Overview[2,1,nnn] = Resultsheet2.cell_value(mci +2 + nnn,12)
        GHG_Table_Overview[2,2,nnn] = Resultsheet2.cell_value(mci +2 + nnn,22)
        GHG_Table_Overview[2,3,nnn] = Resultsheet2.cell_value(mci +2 + nnn,32)
        GHG_Table_Overview[2,4,nnn] = Resultsheet2.cell_value(mci +2 + nnn,42)
        GHG_Table_Overview[2,5,nnn] = Resultsheet2.cell_value(mci +2 + nnn,52)    
                
    
    # Full ME scenario
    Path = os.path.join(RECC_Paths.results_path,ThreeSectoList[-1],'SysVar_TotalGHGFootprint.xls')
    Resultfile   = xlrd.open_workbook(Path)
    Resultsheet1 = Resultfile.sheet_by_name('Cover')
    UUID         = Resultsheet1.cell_value(3,2)
    Resultfile2  = xlrd.open_workbook(os.path.join(RECC_Paths.results_path,ThreeSectoList[-1],'ODYM_RECC_ModelResults_' + UUID + '.xlsx'))
    Resultsheet2 = Resultfile2.sheet_by_name('Model_Results')
    # Find the index for the recycling credit and others:
    tci = 1
    while True:
        if Resultsheet2.cell_value(tci, 0) == 'GHG emissions, system-wide _3579di':
            break # that gives us the right index to read the recycling credit from the result table.
        tci += 1
    
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
        
    for nnn in range(0,2):        
        GHG_Table_Overview[1,0,nnn] = Resultsheet2.cell_value(tci +2 + nnn,8)
        GHG_Table_Overview[1,1,nnn] = Resultsheet2.cell_value(tci +2 + nnn,12)
        GHG_Table_Overview[1,2,nnn] = Resultsheet2.cell_value(tci +2 + nnn,22)
        GHG_Table_Overview[1,3,nnn] = Resultsheet2.cell_value(tci +2 + nnn,32)
        GHG_Table_Overview[1,4,nnn] = Resultsheet2.cell_value(tci +2 + nnn,42)
        GHG_Table_Overview[1,5,nnn] = Resultsheet2.cell_value(tci +2 + nnn,52)
        
        GHG_Table_Overview[3,0,nnn] = Resultsheet2.cell_value(mci +2 + nnn,8)
        GHG_Table_Overview[3,1,nnn] = Resultsheet2.cell_value(mci +2 + nnn,12)
        GHG_Table_Overview[3,2,nnn] = Resultsheet2.cell_value(mci +2 + nnn,22)
        GHG_Table_Overview[3,3,nnn] = Resultsheet2.cell_value(mci +2 + nnn,32)
        GHG_Table_Overview[3,4,nnn] = Resultsheet2.cell_value(mci +2 + nnn,42)
        GHG_Table_Overview[3,5,nnn] = Resultsheet2.cell_value(mci +2 + nnn,52)      
        

    return GHG_Table_Overview

# code for script to be run as standalone function
if __name__ == "__main__":
    main()
