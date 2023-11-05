# -*- coding: utf-8 -*-
"""
Created on Thu Sep 28 13:54:19 2023

@author: spauliuk

This script takes a number of RECC scenarios (as defined in list), 
loads a number of results and then compiles selected results 
into different visualisations of the energy service cascade.

Works together with control workbook
RECCv2.5_EXPORT_Combine_Select.xlxs

Documentation and how to in RECCv2.5_EXPORT_Combine_Select.xlxs
"""

import os
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
import numpy as np
import RECC_Paths # Import path file

def get_esc_data_from_pandas(ps,selectI,selectR,cscenss):
    # return numpay array with esc data for given indicator, region, and scenario list
    pst     = ps[ps['Indicator'].isin([selectI]) & ps['Region'].isin(selectR) & ps['Scenario'].isin(cscenss)] # Select the specified data and transpose them for plotting
    unit    = pst.iloc[0]['Unit']
    Data    = pst.drop(['Scenario','Indicator', 'Region', 'Sectors', 'Unit'], axis=1).values
    return Data, unit

# Definitions/Specifications
CP            = os.path.join(RECC_Paths.results_path,'RECCv2.5_EXPORT_Combine_Select.xlsx')   
CF            = openpyxl.load_workbook(CP)
CS            = CF['Cover'].cell(4,4).value
outpath       = CF[CS].cell(1,2).value
fn_add        = CF[CS].cell(1,4).value


scen = [] # list of target scenarios

# Read scenarios
r = 4
while True:
    if CF[CS].cell(r,1).value is None:
        break
    if CF[CS].cell(r,1).value == 1: # Only if the SELECT flag in col. A is set to 1.
        scen.append(CF[CS].cell(r,2).value)
    r += 1

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


# open data file with results
fn = os.path.join(RECC_Paths.export_path,outpath,'Results_Extracted_RECCv2.5_' + fn_add + '_sep.xlsx')
ps = pd.read_excel(fn, sheet_name='Results', index_col=0) # plot sheet

# determine ESC indicators and plot ESC cascades
# find ESC parameters in Res array: outer index: Ar, middle index: ti, inner indicator: scen

for c in range(0,len(ctitles)):
    if ctypes[c] == 'version_1':
        # get scenario list and length
        if cscens[c] == 'All':
            cscenss = scen
        else:
            cscenss = cscens[c].split(';')
        nocs = len(cscenss)
        selectR = [cregs[c]]
        
        # Define data container
        esc_data = np.zeros((6,46,nocs)) # 6 decoupling indices, 46 years, nocs scenarios
        
        # EXTRACT data and convert to ESC data array
        # Decoupling 1: Lower stock levels
        # Decoupling 2: Operational energy per stock, etc.
        Data_edx,  unit = get_esc_data_from_pandas(ps,'Energy cons., use phase, res+non-res buildings',selectR,cscenss)
        Data_rebx, unit = get_esc_data_from_pandas(ps,'In-use stock, res. buildings',selectR,cscenss)
        Data_nrbx, unit = get_esc_data_from_pandas(ps,'In-use stock, nonres. buildings',selectR,cscenss)
        # Decoupling 3: Build-up material per stock
        Data_matm, unit = get_esc_data_from_pandas(ps,'Final consumption of materials',selectR,cscenss)
        # Decoupling 4: circular material use rate
        Data_rec1, unit = get_esc_data_from_pandas(ps,'ReUse of materials in products, construction grade steel',selectR,cscenss)
        Data_rec2, unit = get_esc_data_from_pandas(ps,'ReUse of materials in products, concrete',selectR,cscenss)
        Data_rec3, unit = get_esc_data_from_pandas(ps,'ReUse of materials in products, wood and wood products',selectR,cscenss)
        Data_rec4, unit = get_esc_data_from_pandas(ps,'Secondary construction steel',selectR,cscenss)
        # Decoupling 6: GHG per energy use
        Data_ghg1, unit = get_esc_data_from_pandas(ps,'GHG emissions, res. buildings, use phase',selectR,cscenss)
        Data_ghg2, unit = get_esc_data_from_pandas(ps,'GHG emissions, non-res. buildings, use phase',selectR,cscenss)
        Data_ghg3, unit = get_esc_data_from_pandas(ps,'GHG emissions, res+non-res buildings, energy supply',selectR,cscenss)
        # Decoupling 5: material footprint per final consumption
        Data_maf1, unit = get_esc_data_from_pandas(ps,'Material footprint, metal ores, system-wide',selectR,cscenss)
        Data_maf2, unit = get_esc_data_from_pandas(ps,'Material footprint, non-metallic minerals, system-wide',selectR,cscenss)
        Data_maf3, unit = get_esc_data_from_pandas(ps,'Material footprint, biomass (dry weight), system-wide',selectR,cscenss)

        esc_data[0,:,:] = (Data_rebx + Data_nrbx).transpose()
        esc_data[1,:,:] = (Data_edx  / (Data_rebx + Data_nrbx)).transpose()
        esc_data[2,:,:] = (Data_matm / (Data_rebx + Data_nrbx)).transpose() 
        esc_data[3,:,:] = ((Data_rec1 + Data_rec2 + Data_rec3 + Data_rec4)/Data_matm).transpose()
        esc_data[4,:,:] = ((Data_maf1 + Data_maf2 + Data_maf3)/Data_matm).transpose()
        esc_data[5,:,:] = ((Data_ghg1 + Data_ghg2 + Data_ghg3)/Data_edx).transpose()
        
        # normalize esc data so that all time series are in relation to SSP2 (maximum):
        for mm in range(0,46):
            esc_data[0,mm,:] = esc_data[0,mm,:] / esc_data[0,mm,:].max()

        # normalize the esc_data
        esc_data[:,0,:] = 0 # start year is 2016, 2015 data are not considered.
        esc_data_Divisor = np.einsum('cs,t->cts',esc_data[:,1,:],np.ones(46))
        esc_data_n = np.divide(esc_data, esc_data_Divisor, out=np.zeros_like(esc_data_Divisor), where=esc_data_Divisor!=0)
        
        # Plot results
        fig, axs = plt.subplots(nrows=1, ncols=6 , figsize=(21, 3))        
        fig.suptitle('Energy service cascade, ' + cregs[c],fontsize=18)
        ProxyHandlesList = []   # For legend 
        
        axs[0].plot(np.arange(2016,2061), esc_data_n[0,1::,:], linewidth = 1.3)
        plta = Line2D(np.arange(2016,2061), esc_data_n[0,1::,:], linewidth = 1.3)
        ProxyHandlesList.append(plta) # create proxy artist for legend    
        axs[0].set_title('(1) Stock per service')
        
        axs[1].plot(np.arange(2016,2061), esc_data_n[1,1::,:], linewidth = 1.3)
        plta = Line2D(np.arange(2016,2061), esc_data_n[1,1::,:], linewidth = 1.3)
        ProxyHandlesList.append(plta) # create proxy artist for legend    
        axs[1].set_title('(2) Operational energy per stock')
        
        axs[2].plot(np.arange(2016,2061), esc_data_n[2,1::,:], linewidth = 1.3)
        plta = Line2D(np.arange(2016,2061), esc_data_n[2,1::,:], linewidth = 1.3)
        ProxyHandlesList.append(plta) # create proxy artist for legend    
        axs[2].set_title('(3) Build-up material per stock')
        
        axs[3].plot(np.arange(2016,2061), esc_data_n[3,1::,:], linewidth = 1.3)
        plta = Line2D(np.arange(2016,2061), esc_data_n[3,1::,:], linewidth = 1.3)
        ProxyHandlesList.append(plta) # create proxy artist for legend    
        axs[3].set_title('(4) Circlar material use rate')
        
        axs[4].plot(np.arange(2016,2061), esc_data_n[4,1::,:], linewidth = 1.3)
        plta = Line2D(np.arange(2016,2061), esc_data_n[4,1::,:], linewidth = 1.3)
        ProxyHandlesList.append(plta) # create proxy artist for legend    
        axs[4].set_title('(5) material footprint per final consumption')        
        
        axs[5].plot(np.arange(2016,2061), esc_data_n[5,1::,:], linewidth = 1.3)
        plta = Line2D(np.arange(2016,2061), esc_data_n[5,1::,:], linewidth = 1.3)
        ProxyHandlesList.append(plta) # create proxy artist for legend    
        axs[5].set_title('(6) GHG emissions per energy use')
        
        Labels = cscenss
        
        fig.legend(Labels, shadow = False, prop={'size':14},ncol=1, loc = 'upper center',bbox_to_anchor=(0.5, -0.02)) 
        plt.tight_layout()
        plt.show()
        title = ctitles[c]
        fig.savefig(os.path.join(os.path.join(RECC_Paths.export_path,outpath), title + '.png'), dpi=150, bbox_inches='tight')



#
#
#
# The end.
#
#
#