# -*- coding: utf-8 -*-
"""
Created on Thu Sep 28 13:54:19 2023

@author: spauliuk
"""
import os
import plotnine
import openpyxl
from plotnine import *
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
import numpy as np
import RECC_Paths # Import path file


def get_RECC_resfile_pos(Label,region,Resultsheet):
    # Find the index for the given Label
    idx = 1    
    while True:
        if Resultsheet.cell(idx, 1).value == Label and Resultsheet.cell(idx, 3).value == region:
            break # that gives us the right index to read the Label from the result table.
        idx += 1 
    return idx


CP            = os.path.join(RECC_Paths.results_path,'RECCv2.5_EXPORT_Combine_Select_2.xlsx')   
CF            = openpyxl.load_workbook(CP)
CS            = CF['Cover'].cell(4,4).value
outpath       = CF[CS].cell(1,6).value
fn_add        = CF[CS].cell(1,8).value

# Definitions/Specifications
Model_id      = CF[CS].cell(1,2).value
Model_date    = CF[CS].cell(1,4).value
outpath       = CF[CS].cell(1,6).value
fn_add        = CF[CS].cell(1,8).value
glob_agg      = CF[CS].cell(1,10).value

# Read specs and variable matchings
scen = [] # list of target scenarios
offs = [] # list of offsets for the different scenarios rel. to starting position
secs = [] # list of matching results folders for the different scenarios

# Read scenarios
r = 4
while True:
    if CF[CS].cell(r,1).value is None:
        break
    if CF[CS].cell(r,1).value == 1: # Only if the SELECT flag in col. A is set to 1.
        scen.append(CF[CS].cell(r,2).value)
        offs.append(CF[CS].cell(r,5).value)
        secs.append([])
        for sl in range(0,20):
            secs[-1].append(CF[CS].cell(r,6+sl).value)
    r += 1
    
# Move to parameter list:
while True:
    if CF[CS].cell(r,2).value == 'Target indicator label':
        break    
    r += 1
r += 1    

# Read indicator list:
ti = []    # target indicator
ri = []    # RECC indicator
tu = []    # target unit
ru = []    # RECC unit
cf = []    # Conv. factor
sl = []    # sector list
rr = []    # RECC regional resolution

while True:
    if CF[CS].cell(r,2).value is None:
        break
    ti.append(CF[CS].cell(r,2).value)
    ri.append(CF[CS].cell(r,3).value)
    tu.append(CF[CS].cell(r,4).value)
    ru.append(CF[CS].cell(r,5).value)
    cf.append(CF[CS].cell(r,6).value)
    sl.append(CF[CS].cell(r,7).value)
    rr.append(CF[CS].cell(r,8).value)
    r += 1

nos = len(scen) # number of scenarios
noi = len(ti)   # number of indicators

# split sector labels at '+' for different sectors
sL = [i.split('+') for i in sl]

# Move to region aggregation definition:
while True:
    if CF[CS].cell(r,2).value == 'AggRegion':
        break    
    r += 1
r += 1   

# Region List
Ar = []    # Aggregated (target) regions for export
Dr = []    # RECC detailed regions

r0 = r

while True:
    if CF[CS].cell(r,2).value is None:
        break
    Ar.append(CF[CS].cell(r,2).value)
    c = 3
    Dr.append([])
    while True:
        if CF[CS].cell(r,c).value is None:
            break
        Dr[r-r0].append(CF[CS].cell(r,c).value)
        c += 1    
    r += 1

nor = len(Ar)

# iterate over scenarios, parse results:
Res  = np.zeros((nor*nos*noi,46)) # main result array 46 years
ResC = np.zeros((nor*nos*noi,6))  # main results, 6 columns for cumulative results

Folders = list(set([item for sublist in secs for item in sublist]))

# Extract, aggregated, and reformat results
for rf in Folders:
    if rf is not None: # if a folder is given, extract all indicators and write to corresponding position in result array:
        print('Reading data from ' + rf)
        RECC_ResFile = [filename for filename in os.listdir(os.path.join(RECC_Paths.results_path,rf)) if filename.startswith('ODYM_RECC_ModelResults_')]
        RECC_RF      = openpyxl.load_workbook(os.path.join(RECC_Paths.results_path,rf,RECC_ResFile[0]))
        RECC_RS      = RECC_RF['Model_Results']
        parts = rf.split('__')
        region = parts[0]
        sector = parts[3]
        # look for where to put data from this result folder:
        for s in range(0,nos): # iterate over all selected scenarios
            for i in range(0,noi): # for all indicators
                print('Reading data for ' + ti[i])
                for r in range(0,nor): # for all regions
                    if region == Ar[r] or region in Dr[r]: # the current result file region is or is part of the current target region
                        targetpos = r*nos*noi + i * nos + s # position in Res array: outer index: region, middle index: indicator, inner index: scenario
                        if sector in sL[i]: # if current sector is part of target sector for indicator
                            if rf in secs[s]: # if current folder is in list for currect scenario --> extract results!
                                if rr[i] == 'Aggregate': # use indicator for aggregate region label and add to results:
                                    idx = get_RECC_resfile_pos(ri[i],region,RECC_RS) # position of that indicator for that region in the RECC result file
                                    for t in range(0,46): # read and add values
                                        Res[targetpos,t] += RECC_RS.cell(idx+offs[s],t+8).value * cf[i] # Value from Excel to array
                                if rr[i] == 'Aggregate_from_Single': # use indicator for aggregate region label and add to results:
                                    for sir in range(0,len(Dr[r])):
                                        idx = get_RECC_resfile_pos(ri[i],Dr[r][sir],RECC_RS) # position of that indicator for that region in the RECC result file
                                        for t in range(0,46): # read and add values
                                            Res[targetpos,t] += RECC_RS.cell(idx+offs[s],t+8).value * cf[i] # Value from Excel to array

# Special: All regions to global aggregate, if region is outer index and all regions add up
if glob_agg    == 'True':
    region_no  = nos*noi
    Res_r      = Res.reshape((nor,nos*noi,46)) # reshape to region as separate dimension
    Res_r_agg  = Res_r.sum(axis=0) # sum over regions
    Res_rC     = ResC.reshape((nor,nos*noi,6)) # reshape to region as separate dimension
    Res_rC_agg = Res_rC.sum(axis=0) # sum over regions

r = 1
# Move to parameter list:
while True:
    if CF[CS].cell(r,1).value == 'Define ESC plot':
        break    
    r += 1
r += 1

ctitles = []
ctypes  = []
cregs   = []
cscens  = []

while True:
    if CF[CS].cell(r,2).value is None:
        break    
    ctitles.append(CF[CS].cell(r,2).value)
    ctypes.append(CF[CS].cell(r,3).value)
    cregs.append(CF[CS].cell(r,4).value)
    cscens.append(CF[CS].cell(r,5).value)
    r += 1

# determine ESC indicators and plot ESC cascades
# find ESC parameters in Res array: outer index: Ar, middle index: ti, inner indicator: scen

for c in range(0,len(ctitles)):
    if ctypes[c] == 'version_1':
        if cscens[c] == 'All':
            nocs = nos # number of cascade scenarios nocs = number of scenarios nos
        else:
            nocs = len(cscens[c].split(';'))
        if cregs[c] == 'Global':
            rind = -1
        else:
            rind = Ar.index(cregs[c])
        # Define data container
        esc_data = np.zeros((6,46,nocs)) # 6 decoupling indices, 46 years, nocs scenarios
        
        # Decoupling 2: Operational energy per stock:
        edx  = ti.index('Energy cons., use phase, res+non-res buildings')
        rebx = ti.index('In-use stock, res. buildings')
        nrbx = ti.index('In-use stock, nonres. buildings')
        for sc in range(0,nocs):
            targetpos_edx  = rind*nos*noi + edx  * nos + scen.index(cscens[c].split(';')[sc]) # position in Res array: outer index: region, middle index: indicator, inner index: scenario
            targetpos_rebx = rind*nos*noi + rebx * nos + scen.index(cscens[c].split(';')[sc])
            targetpos_nrbx = rind*nos*noi + nrbx * nos + scen.index(cscens[c].split(';')[sc])
            esc_data[1,:,sc] = Res[targetpos_edx,:] / (Res[targetpos_rebx,:] + Res[targetpos_nrbx,:])           


        # normalize the esc_data
        esc_data_Divisor = np.einsum('cs,t->cts',esc_data[:,0,:],np.ones(46))
        esc_data_n = np.divide(esc_data, esc_data_Divisor, out=np.zeros_like(esc_data_Divisor), where=esc_data_Divisor!=0)
        
        fig, axs = plt.subplots(nrows=1, ncols=6 , figsize=(21, 3))
        fig.suptitle('Energy service cascade.',fontsize=18)
        
        ProxyHandlesList = []   # For legend 
        
        axs[4].plot(np.arange(2016,2061), esc_data_n[1,1::,:], linewidth = 1.3)
        plta = Line2D(np.arange(2016,2061), esc_data_n[1,1::,:], linewidth = 1.3)
        ProxyHandlesList.append(plta) # create proxy artist for legend    
        
        Labels = cscens[c].split(';')
        
        fig.legend(Labels, shadow = False, prop={'size':14},ncol=1, loc = 'upper center',bbox_to_anchor=(0.5, -0.02)) 
        plt.tight_layout()
        plt.show()
        title = ctitles[c]
        fig.savefig(os.path.join(os.path.join(RECC_Paths.export_path,outpath), title + '.png'), dpi=150, bbox_inches='tight')

# for m in range(0,len(ptitles)):
#     if ptypes[m] == 'line_fixedIndicator_fixedRegion_varScenario':
#     # Plot single indicator for one region and all scenarios
#         selectI = [pinds[m]]
#         selectR = [pregs[m]]
#         if pscens[m] == 'All':
#             pst       = ps[ps['Indicator'].isin(selectI) & ps['Region'].isin(selectR)].T # Select the specified data and transpose them for plotting
#             title_add = '_all_scenarios'
#         else:
#             selectS = pscens[m].split(';')
#             pst     = ps[ps['Indicator'].isin(selectI) & ps['Region'].isin(selectR) & ps['Scenario'].isin(selectS)].T # Select the specified data and transpose them for plotting
#             title_add = '_select_scenarios_' + str(len(selectS))
#         pst.columns = pst.iloc[2] # Set scenario column (with unique labels) as column names
#         unit    = pst.iloc[4][1] 
#         pst.drop(['Region','Indicator','Scenario','Sectors','Unit'], inplace=True) # Delete labels that are not needed
#         pst.plot(kind = 'line', figsize=(10,5), ) # plot data, configure plot, and save results
#         plt.xlabel('Year')
#         plt.ylabel(unit)
#         title = ptitles[m] + '_' + selectR[0] + title_add
#         plt.title(title)
#         plt.savefig(os.path.join(os.path.join(RECC_Paths.export_path,outpath), title + '.png'), dpi=150, bbox_inches='tight')

