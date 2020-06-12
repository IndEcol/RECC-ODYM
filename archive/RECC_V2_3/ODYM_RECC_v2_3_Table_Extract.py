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
    

    
    GHG_TableX       = np.zeros((4,6)) # Table X format

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
        
    GHG_TableX[0,0] = Resultsheet2.cell_value(tci+ 3,8)
    GHG_TableX[0,1] = Resultsheet2.cell_value(tci+ 3,12)
    GHG_TableX[0,2] = Resultsheet2.cell_value(tci+ 3,22)
    GHG_TableX[0,3] = Resultsheet2.cell_value(tci+ 3,32)
    GHG_TableX[0,4] = Resultsheet2.cell_value(tci+ 3,42)
    GHG_TableX[0,5] = Resultsheet2.cell_value(tci+ 3,52)
    
    GHG_TableX[2,0] = Resultsheet2.cell_value(mci+ 3,8)
    GHG_TableX[2,1] = Resultsheet2.cell_value(mci+ 3,12)
    GHG_TableX[2,2] = Resultsheet2.cell_value(mci+ 3,22)
    GHG_TableX[2,3] = Resultsheet2.cell_value(mci+ 3,32)
    GHG_TableX[2,4] = Resultsheet2.cell_value(mci+ 3,42)
    GHG_TableX[2,5] = Resultsheet2.cell_value(mci+ 3,52)    
                
    
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
        
    GHG_TableX[1,0] = Resultsheet2.cell_value(tci+ 3,8)
    GHG_TableX[1,1] = Resultsheet2.cell_value(tci+ 3,12)
    GHG_TableX[1,2] = Resultsheet2.cell_value(tci+ 3,22)
    GHG_TableX[1,3] = Resultsheet2.cell_value(tci+ 3,32)
    GHG_TableX[1,4] = Resultsheet2.cell_value(tci+ 3,42)
    GHG_TableX[1,5] = Resultsheet2.cell_value(tci+ 3,52)
    
    GHG_TableX[3,0] = Resultsheet2.cell_value(mci+ 3,8)
    GHG_TableX[3,1] = Resultsheet2.cell_value(mci+ 3,12)
    GHG_TableX[3,2] = Resultsheet2.cell_value(mci+ 3,22)
    GHG_TableX[3,3] = Resultsheet2.cell_value(mci+ 3,32)
    GHG_TableX[3,4] = Resultsheet2.cell_value(mci+ 3,42)
    GHG_TableX[3,5] = Resultsheet2.cell_value(mci+ 3,52)      
    

    
    return GHG_TableX

# code for script to be run as standalone function
if __name__ == "__main__":
    main()
