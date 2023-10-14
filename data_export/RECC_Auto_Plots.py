# -*- coding: utf-8 -*-
"""
Created on Thu Sep 28 13:54:19 2023

@author: spauliuk

This script takes a number of RECC scenarios (as defined in list), 
loads a number of results and then compiles selected results 
into different visualisations, like line plots and bar charts.

Works together with control workbook
RECCv2.5_EXPORT_Combine_Select.xlxs

Documentation and how to in RECCv2.5_EXPORT_Combine_Select.xlxs

"""
import os
import plotnine
import openpyxl
from plotnine import *
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import RECC_Paths # Import path file


CP            = os.path.join(RECC_Paths.results_path,'RECCv2.5_EXPORT_Combine_Select.xlsx')   
CF            = openpyxl.load_workbook(CP)
CS            = CF['Cover'].cell(4,4).value
outpath       = CF[CS].cell(1,2).value
fn_add        = CF[CS].cell(1,4).value

r = 1
# Move to parameter list:
while True:
    if CF[CS].cell(r,1).value == 'Define Plots':
        break    
    r += 1
r += 1

ptitles = []
ptypes  = []
pinds   = []
pregs   = []
pscens  = []
prange  = []
while True:
    if CF[CS].cell(r,2).value is None:
        break    
    ptitles.append(CF[CS].cell(r,2).value)
    ptypes.append(CF[CS].cell(r,3).value)
    pinds.append(CF[CS].cell(r,4).value)
    pregs.append(CF[CS].cell(r,5).value)
    pscens.append(CF[CS].cell(r,6).value)
    prange.append(CF[CS].cell(r,7).value)
    r += 1

# open data file with results
fn = os.path.join(RECC_Paths.export_path,outpath,'Results_Extracted_RECCv2.5_' + fn_add + '_sep.xlsx')
ps = pd.read_excel(fn, sheet_name='Results', index_col=0) # plot sheet
pc = pd.read_excel(fn, sheet_name='Results_Cumulative', index_col=0) # plot sheet cumulative


for m in range(0,len(ptitles)):
    if ptypes[m] == 'line_fixedIndicator_fixedRegion_varScenario':
    # Plot single indicator for one region and all scenarios
        selectI = [pinds[m]]
        selectR = [pregs[m]]
        if pscens[m] == 'All':
            pst       = ps[ps['Indicator'].isin(selectI) & ps['Region'].isin(selectR)].T # Select the specified data and transpose them for plotting
            title_add = '_all_scenarios'
        else:
            selectS = pscens[m].split(';')
            pst     = ps[ps['Indicator'].isin(selectI) & ps['Region'].isin(selectR) & ps['Scenario'].isin(selectS)].T # Select the specified data and transpose them for plotting
            title_add = '_select_scenarios_' + str(len(selectS))
        pst.columns = pst.iloc[2] # Set scenario column (with unique labels) as column names
        unit    = pst.iloc[4][1] 
        pst.drop(['Region','Indicator','Scenario','Sectors','Unit'], inplace=True) # Delete labels that are not needed
        pst.plot(kind = 'line', figsize=(10,5), ) # plot data, configure plot, and save results
        plt.xlabel('Year')
        plt.ylabel(unit)
        title = ptitles[m] + '_' + selectR[0] + title_add
        plt.title(title)
        plt.savefig(os.path.join(os.path.join(RECC_Paths.export_path,outpath), title + '.png'), dpi=150, bbox_inches='tight')

    if ptypes[m] == 'line_fixedIndicator_varRegion_fixedScenario':
    # Plot single indicator for one scenario and all regions
        selectI = [pinds[m]]
        selectS = [pscens[m]]
        if pregs[m] == 'All':
            pst    = ps[ps['Indicator'].isin(selectI) & ps['Scenario'].isin(selectS)].T # Select the specified data and transpose them for plotting
            title_add = '_all_regions'
        else:
            selectR = pregs[m].split(';')
            pst    = ps[ps['Indicator'].isin(selectI) & ps['Region'].isin(selectR) & ps['Scenario'].isin(selectS)].T # Select the specified data and transpose them for plotting
            title_add = '_select_regions_' + str(len(selectS))
        pst.columns = pst.iloc[0] # Set scenario column (with unique labels) as column names
        unit    = pst.iloc[4][1] 
        pst.drop(['Region','Indicator','Scenario','Sectors','Unit'], inplace=True) # Delete labels that are not needed
        pst.plot(kind = 'line', figsize=(10,5), ) # plot data, configure plot, and save results
        plt.xlabel('Year')
        plt.ylabel(unit)
        title = ptitles[m] + '_' + selectS[0] + title_add
        plt.title(title)
        plt.savefig(os.path.join(os.path.join(RECC_Paths.export_path,outpath), title +'.png'), dpi=150, bbox_inches='tight')

    if ptypes[m] == 'hbar_cum_fixedIndicator_fixedRegion_varScenario':
        # Plot bar graph with cumulative indicator by scenario
        selectI = [pinds[m]]
        selectR = [pregs[m]]
        if pscens[m] == 'All':
            pst    = pc[pc['Indicator'].isin(selectI) & pc['Region'].isin(selectR)] # Select the specified data and transpose them for plotting
            title_add = '_all_scenarios'
        else:
            selectS = pscens[m].split(';')
            pst    = pc[pc['Indicator'].isin(selectI) & pc['Region'].isin(selectR) & ps['Scenario'].isin(selectS)] # Select the specified data and transpose them for plotting
            title_add = '_select_scenarios_' + str(len(selectS))
        pst.set_index('Scenario', inplace=True)
        pst[prange[m]]
        unit = pst.iloc[0][3]
        pst.plot.barh(y=prange[m])
        plt.xlabel(unit)
        title = ptitles[m] + '_' + selectR[0] + title_add
        plt.title(title)
        plt.savefig(os.path.join(os.path.join(RECC_Paths.export_path,outpath), title +'.png'), dpi=150, bbox_inches='tight')
        
    if ptypes[m] == 'hbar_cum_fixedIndicator_varRegion_fixedScenario':
        # Plot bar graph with cumulative indicator by scenario
        selectI = [pinds[m]]
        selectS = [pscens[m]]
        if pregs[m] == 'All':
            pst    = pc[pc['Indicator'].isin(selectI) & pc['Scenario'].isin(selectS)] # Select the specified data and transpose them for plotting
            title_add = '_all_regions'
        else:
            selectR = pregs[m].split(';')
            pst    = pc[pc['Indicator'].isin(selectI) & pc['Region'].isin(selectR) & ps['Scenario'].isin(selectS)] # Select the specified data and transpose them for plotting
            title_add = '_select_regions_' + str(len(selectR))
        pst.set_index('Region', inplace=True)
        pst[prange[m]]
        unit = pst.iloc[0][3]
        pst.plot.barh(y=prange[m])
        plt.xlabel(unit)
        title = ptitles[m] + '_' + selectR[0] + title_add
        plt.title(title)
        plt.savefig(os.path.join(os.path.join(RECC_Paths.export_path,outpath), title +'.png'), dpi=150, bbox_inches='tight')
        






# pst.rename(columns={'Scenario': 'Year'}, inplace=True)


# .drop(['Region'])
# pst.insert(0, "index", np.arange(0,51))

# .set_index(np.arange(0,52)) 
# pst.drop([0,1,2,4,5])
# pst.plot(x='variable', y='value')

# ps_2 = pC[pC['Variable'].isin(selectV) & pC['Region'].isin(selectR)]

# ggplot(ps_1p, aes(x='variable', y='value')) + geom_point()
# ggplot(ps_1p, aes(x='variable', y='value',color='Scenario')) + geom_point()
# ggsave(ggplot(ps_1p, aes(x='variable', y='value',color='Scenario')) + geom_point(), filename="plot1.png", device='png', dpi=300, height=25, width=25, )
# ggsave(ggplot(ps_1p, aes(x='variable', y='value', group='Scenario', color='Scenario')) 
#        + geom_line() + geom_point() + theme_classic() 
#        + scale_x_continuous(breaks=range(2020,2060,10) + xlim(2015,2060))
#        + labs(title=selectV[0] + ', ' + unit,x='Year', y = unit), filename="plot1.png", path = pa, dpi=300, height=8, width=12, )
  

# plotd  = pl.plot(ps1.to_numpy()[:,5::].transpose())

# unit    = ps_1['Unit'].iloc[0]
# ps1p   = pd.melt(ps1, id_vars=['Scenario'], value_vars=[i for i in range(2015,2061)])

# psn    = ps[ps['Variable'].isin(selectV) & ps['Region'].isin(selectR)].to_numpy()

# ps1    = ps[ps['Variable'].isin(selectV) & ps['Region'].isin(selectR)]
        
# # ggtitle(selectV[0] + ', ' + unit)        