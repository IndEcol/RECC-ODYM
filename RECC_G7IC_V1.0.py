# -*- coding: utf-8 -*-
"""
Created on July 22, 2018

@authors: spauliuk
"""

"""
File RECC_G7_V1.0.py

Contains the model instructions for the resource efficiency climate change project developed using ODYM: ODYM-RECC

dependencies:
    numpy >= 1.9
    scipy >= 0.14

"""
# Import required libraries:
import os
import sys
import logging as log
import xlrd, xlwt
import numpy as np
import time
import datetime
import scipy.io
import pandas as pd
import shutil   
import uuid
import matplotlib.pyplot as plt   
from matplotlib.lines import Line2D
import importlib
import getpass
from copy import deepcopy
from tqdm import tqdm
from scipy.interpolate import interp1d
import pylab
import pickle

import RECC_Paths # Import path file


#import re
__version__ = str('1.0')
##################################
#    Section 1)  Initialize      #
##################################
# add ODYM module directory to system path
sys.path.insert(0, os.path.join(os.path.join(RECC_Paths.odym_path,'odym'),'modules'))
### 1.1.) Read main script parameters
# Mylog.info('### 1.1 - Read main script parameters')
ProjectSpecs_Name_ConFile = 'RECC_Config.xlsx'
Model_Configfile = xlrd.open_workbook(ProjectSpecs_Name_ConFile)
ScriptConfig = {'Model Setting': Model_Configfile.sheet_by_name('Config').cell_value(3,3)}
Model_Configsheet = Model_Configfile.sheet_by_name(ScriptConfig['Model Setting'])
#Read debug modus:   
DebugCounter = 0
while Model_Configsheet.cell_value(DebugCounter, 2) != 'Logging_Verbosity':
    DebugCounter += 1
ScriptConfig['Logging_Verbosity'] = Model_Configsheet.cell_value(DebugCounter,3) # Read loggin verbosity once entry was reached.    
# Extract user name from main file
ProjectSpecs_User_Name     = getpass.getuser()

# import packages whose location is now on the system path:    
import ODYM_Classes as msc # import the ODYM class file
importlib.reload(msc)
import ODYM_Functions as msf  # import the ODYM function file
importlib.reload(msf)
import dynamic_stock_model as dsm # import the dynamic stock model library
importlib.reload(dsm)

Name_Script        = Model_Configsheet.cell_value(5,3)
if Name_Script != 'RECC_G7IC_V1.0':  # Name of this script must equal the specified name in the Excel config file
    # TODO: This does not work because the logger was not yet initialized
    # log.critical("The name of the current script '%s' does not match to the sript name specfied in the project configuration file '%s'. Exiting the script.",
    #              Name_Script, 'ODYM_RECC_Test1')
    raise AssertionError('Fatal: The name of the current script does not match to the sript name specfied in the project configuration file. Exiting the script.')
# the model will terminate if the name of the script that is run is not identical to the script name specified in the config file.
Name_Scenario            = Model_Configsheet.cell_value(3,3)
UUID_Scenario            = str(uuid.uuid4())
StartTime                = datetime.datetime.now()
TimeString               = str(StartTime.year) + '_' + str(StartTime.month) + '_' + str(StartTime.day) + '__' + str(StartTime.hour) + '_' + str(StartTime.minute) + '_' + str(StartTime.second)
DateString               = str(StartTime.year) + '_' + str(StartTime.month) + '_' + str(StartTime.day)
ProjectSpecs_Path_Result = os.path.join(RECC_Paths.results_path, Name_Scenario + '_' + TimeString )

if not os.path.exists(ProjectSpecs_Path_Result): # Create model run results directory.
    os.makedirs(ProjectSpecs_Path_Result)
# Initialize logger
if ScriptConfig['Logging_Verbosity'] == 'DEBUG':
    log_verbosity = eval("log.DEBUG")  
log_filename = Name_Scenario + '_' + TimeString + '.md'
[Mylog, console_log, file_log] = msf.function_logger(log_filename, ProjectSpecs_Path_Result,
                                                     log_verbosity, log_verbosity)
# log header and general information
Time_Start = time.time()
ScriptConfig['Current_UUID'] = str(uuid.uuid4())
Mylog.info('# Simulation from ' + time.asctime())
Mylog.info('Unique ID of scenario run: ' + ScriptConfig['Current_UUID'])

### 1.2) Read model control parameters
Mylog.info('### 1.2 - Read model control parameters')
#Read control and selection parameters into dictionary
SCix = 0
# search for script config list entry
while Model_Configsheet.cell_value(SCix, 1) != 'General Info':
    SCix += 1
        
SCix += 2  # start on first data row
while len(Model_Configsheet.cell_value(SCix, 3)) > 0:
    ScriptConfig[Model_Configsheet.cell_value(SCix, 2)] = Model_Configsheet.cell_value(SCix,3)
    SCix += 1

SCix = 0
# search for script config list entry
while Model_Configsheet.cell_value(SCix, 1) != 'Software version selection':
    SCix += 1
        
SCix += 2 # start on first data row
while len(Model_Configsheet.cell_value(SCix, 3)) > 0:
    ScriptConfig[Model_Configsheet.cell_value(SCix, 2)] = Model_Configsheet.cell_value(SCix,3)
    SCix += 1  

Mylog.info('Script: ' + Name_Script + '.py')
Mylog.info('Model script version: ' + __version__)
Mylog.info('Model functions version: ' + msf.__version__())
Mylog.info('Model classes version: ' + msc.__version__())
Mylog.info('Current User: ' + ProjectSpecs_User_Name)
Mylog.info('Current Scenario: ' + Name_Scenario)
Mylog.info(ScriptConfig['Description'])
Mylog.debug('----\n')

### 1.3) Organize model output folder and logger
Mylog.info('### 1.3 Organize model output folder and logger')
#Copy Config file and model script into that folder
shutil.copy(ProjectSpecs_Name_ConFile, os.path.join(ProjectSpecs_Path_Result, ProjectSpecs_Name_ConFile))
shutil.copy(Name_Script + '.py'      , os.path.join(ProjectSpecs_Path_Result, Name_Script + '.py'))

#####################################################
#     Section 2) Read classifications and data      #
#####################################################
Mylog.info('## 2 - Read classification items and define all classifications')
### 2.1) # Read model run config data
Mylog.info('### 2.1 - Read model run config data')
# Note: This part reads the items directly from the Exel master,
# will be replaced by reading them from version-managed csv file.
class_filename = str(ScriptConfig['Version of master classification']) + '.xlsx'
Classfile = xlrd.open_workbook(os.path.join(RECC_Paths.data_path,class_filename))
Classsheet = Classfile.sheet_by_name('MAIN_Table')
ci = 1  # column index to start with
MasterClassification = {}  # Dict of master classifications
while True:
    TheseItems = []
    ri = 10  # row index to start with
    try: 
        ThisName = Classsheet.cell_value(0,ci)
        ThisDim  = Classsheet.cell_value(1,ci)
        ThisID   = Classsheet.cell_value(3,ci)
        ThisUUID = Classsheet.cell_value(4,ci)
        TheseItems.append(Classsheet.cell_value(ri,ci)) # read the first classification item
    except:
        Mylog.info('End of file or formatting error while reading the classification file in column ' + str(ci) + '. Check if all classifications are present. If yes, you are good to go!')
        break
    while True:
        ri += 1
        try:
            ThisItem = Classsheet.cell_value(ri, ci)
        except:
            break
        if ThisItem is not '':
            TheseItems.append(ThisItem)
    MasterClassification[ThisName] = msc.Classification(Name = ThisName, Dimension = ThisDim, ID = ThisID, UUID = ThisUUID, Items = TheseItems)
    ci += 1
    
Mylog.info('Read index table from model config sheet.')
ITix = 0

# search for index table entry
while True:
    if Model_Configsheet.cell_value(ITix, 1) == 'Index Table':
        break
    else:
        ITix += 1
        
IT_Aspects        = []
IT_Description    = []
IT_Dimension      = []
IT_Classification = []
IT_Selector       = []
IT_IndexLetter    = []
ITix += 2 # start on first data row
while True:
    if len(Model_Configsheet.cell_value(ITix,2)) > 0:
        IT_Aspects.append(Model_Configsheet.cell_value(ITix,2))
        IT_Description.append(Model_Configsheet.cell_value(ITix,3))
        IT_Dimension.append(Model_Configsheet.cell_value(ITix,4))
        IT_Classification.append(Model_Configsheet.cell_value(ITix,5))
        IT_Selector.append(Model_Configsheet.cell_value(ITix,6))
        IT_IndexLetter.append(Model_Configsheet.cell_value(ITix,7))        
        ITix += 1
    else:
        break

Mylog.info('Read parameter list from model config sheet.')
PLix = 0
while True: # search for parameter list entry
    if Model_Configsheet.cell_value(PLix, 1) == 'Model Parameters':
        break
    else:
        PLix += 1
        
PL_Names          = []
PL_Description    = []
PL_Version        = []
PL_IndexStructure = []
PL_IndexMatch     = []
PL_IndexLayer     = []
PLix += 2 # start on first data row
while True:
    if len(Model_Configsheet.cell_value(PLix,2)) > 0:
        PL_Names.append(Model_Configsheet.cell_value(PLix,2))
        PL_Description.append(Model_Configsheet.cell_value(PLix,3))
        PL_Version.append(Model_Configsheet.cell_value(PLix,4))
        PL_IndexStructure.append(Model_Configsheet.cell_value(PLix,5))
        PL_IndexMatch.append(Model_Configsheet.cell_value(PLix,6))
        PL_IndexLayer.append(msf.ListStringToListNumbers(Model_Configsheet.cell_value(PLix,7))) # strip numbers out of list string
        PLix += 1
    else:
        break
    
Mylog.info('Read process list from model config sheet.')
PrLix = 0

# search for process list entry
while True:
    if Model_Configsheet.cell_value(PrLix, 1) == 'Process Group List':
        break
    else:
        PrLix += 1
        
PrL_Number         = []
PrL_Name           = []
PrL_Comment        = []
PrL_Type           = []
PrLix += 2 # start on first data row
while True:
    if Model_Configsheet.cell_value(PrLix,2) != '':
        try:
            PrL_Number.append(int(Model_Configsheet.cell_value(PrLix,2)))
        except:
            PrL_Number.append(Model_Configsheet.cell_value(PrLix,2))
        PrL_Name.append(Model_Configsheet.cell_value(PrLix,3))
        PrL_Type.append(Model_Configsheet.cell_value(PrLix,4))
        PrL_Comment.append(Model_Configsheet.cell_value(PrLix,5))
        PrLix += 1
    else:
        break    

Mylog.info('Read model run control from model config sheet.')
PrLix = 0

# search for model flow control entry
while True:
    if Model_Configsheet.cell_value(PrLix, 1) == 'Model flow control':
        break
    else:
        PrLix += 1

# start on first data row
PrLix += 2
while True:
    if Model_Configsheet.cell_value(PrLix, 2) != '':
        try:
            ScriptConfig[Model_Configsheet.cell_value(PrLix, 2)] = Model_Configsheet.cell_value(PrLix,3)
        except:
            None
        PrLix += 1
    else:
        break  

Mylog.info('Read model output control from model config sheet.')
PrLix = 0

# search for model flow control entry
while True:
    if Model_Configsheet.cell_value(PrLix, 1) == 'Model output control':
        break
    else:
        PrLix += 1

# start on first data row
PrLix += 2
while True:
    if Model_Configsheet.cell_value(PrLix, 2) != '':
        try:
            ScriptConfig[Model_Configsheet.cell_value(PrLix, 2)] = Model_Configsheet.cell_value(PrLix,3)
        except:
            None
        PrLix += 1
    else:
        break  

Mylog.info('Define model classifications and select items for model classifications according to information provided by config file.')
ModelClassification  = {} # Dict of model classifications
for m in range(0,len(IT_Aspects)):
    ModelClassification[IT_Aspects[m]] = deepcopy(MasterClassification[IT_Classification[m]])
    EvalString = msf.EvalItemSelectString(IT_Selector[m],len(ModelClassification[IT_Aspects[m]].Items))
    if EvalString.find(':') > -1: # range of items is taken
        RangeStart = int(EvalString[0:EvalString.find(':')])
        RangeStop  = int(EvalString[EvalString.find(':')+1::])
        ModelClassification[IT_Aspects[m]].Items = ModelClassification[IT_Aspects[m]].Items[RangeStart:RangeStop]           
    elif EvalString.find('[') > -1: # selected items are taken
        ModelClassification[IT_Aspects[m]].Items = [ModelClassification[IT_Aspects[m]].Items[i] for i in eval(EvalString)]
    elif EvalString == 'all':
        None
    else:
        Mylog.error('Item select error for aspect ' + IT_Aspects[m] + ' were found in datafile.')
        break
    
### 2.2) # Define model index table and parameter dictionary
Mylog.info('### 2.2 - Define model index table and parameter dictionary')
Model_Time_Start = int(min(ModelClassification['Time'].Items))
Model_Time_End   = int(max(ModelClassification['Time'].Items))
Model_Duration   = Model_Time_End - Model_Time_Start + 1

Mylog.info('Define index table dataframe.')
IndexTable = pd.DataFrame({'Aspect'        : IT_Aspects,  # 'Time' and 'Element' must be present!
                           'Description'   : IT_Description,
                           'Dimension'     : IT_Dimension,
                           'Classification': [ModelClassification[Aspect] for Aspect in IT_Aspects],
                           'IndexLetter'   : IT_IndexLetter})  # Unique one letter (upper or lower case) indices to be used later for calculations.

# Default indexing of IndexTable, other indices are produced on the fly
IndexTable.set_index('Aspect', inplace=True)

# Add indexSize to IndexTable:
IndexTable['IndexSize'] = pd.Series([len(IndexTable.Classification[i].Items) for i in range(0, len(IndexTable.IndexLetter))],
                                    index=IndexTable.index)

# list of the classifications used for each indexletter
IndexTable_ClassificationNames = [IndexTable.Classification[i].Name for i in range(0, len(IndexTable.IndexLetter))]

# Define shortcuts for the most important index sizes:
Nt = len(IndexTable.Classification[IndexTable.index.get_loc('Time')].Items)
Ne = len(IndexTable.Classification[IndexTable.index.get_loc('Element')].Items)
Nc = len(IndexTable.Classification[IndexTable.index.get_loc('Cohort')].Items)
Nr = len(IndexTable.Classification[IndexTable.index.get_loc('Region')].Items)
NG = len(IndexTable.Classification[IndexTable.set_index('IndexLetter').index.get_loc('G')].Items)
Ng = len(IndexTable.Classification[IndexTable.set_index('IndexLetter').index.get_loc('g')].Items)
NS = len(IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items)
NR = len(IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
Nw = len(IndexTable.Classification[IndexTable.index.get_loc('Waste_Scrap')].Items)
Nm = len(IndexTable.Classification[IndexTable.index.get_loc('Engineering materials')].Items)
Nk = len(IndexTable.Classification[IndexTable.index.get_loc('Component')].Items)
NX = len(IndexTable.Classification[IndexTable.index.get_loc('Extensions')].Items)
Nn = len(IndexTable.Classification[IndexTable.index.get_loc('Energy')].Items)
NV = len(IndexTable.Classification[IndexTable.set_index('IndexLetter').index.get_loc('V')].Items)
#IndexTable.ix['t']['Classification'].Items # get classification items

Mylog.info('Read model data and parameters.')

ParFileName = os.path.join(RECC_Paths.data_path,'RECC_ParameterDict_' + ScriptConfig['Model Setting'] + '_V1.dat')
try: # Load Pickle parameter dict to save processing time
    ParFileObject = open(ParFileName,'rb')  
    ParameterDict = pickle.load(ParFileObject)  
    ParFileObject.close()  
    Mylog.info('Model data and parameters were read from pickled file.')
except:
    ParameterDict = {}
    mo_start = 0 # set mo for re-reading a certain parameter
    for mo in range(mo_start,len(PL_Names)):
        #mo = 30 # set mo for re-reading a certain parameter
        #ParPath = os.path.join(os.path.abspath(os.path.join(ProjectSpecs_Path_Main, '.')), 'ODYM_RECC_Database', PL_Version[mo])
        ParPath = os.path.join(RECC_Paths.data_path, PL_Names[mo] + '_' + PL_Version[mo])
        Mylog.info('Reading parameter ' + PL_Names[mo])
        #MetaData, Values = msf.ReadParameter(ParPath = ParPath,ThisPar = PL_Names[mo], ThisParIx = PL_IndexStructure[mo], IndexMatch = PL_IndexMatch[mo], ThisParLayerSel = PL_IndexLayer[mo], MasterClassification,IndexTable,IndexTable_ClassificationNames,ScriptConfig,Mylog) # Do not change order of parameters handed over to function!
        # Do not change order of parameters handed over to function!
        MetaData, Values = msf.ReadParameterV2(ParPath, PL_Names[mo], PL_IndexStructure[mo], PL_IndexMatch[mo],
                                             PL_IndexLayer[mo], MasterClassification, IndexTable,
                                             IndexTable_ClassificationNames, ScriptConfig, Mylog, False)
        ParameterDict[PL_Names[mo]] = msc.Parameter(Name=MetaData['Dataset_Name'], ID=MetaData['Dataset_ID'],
                                                    UUID=MetaData['Dataset_UUID'], P_Res=None, MetaData=MetaData,
                                                    Indices=PL_IndexStructure[mo], Values=Values, Uncert=None,
                                                    Unit=MetaData['Dataset_Unit'])
        Mylog.info('Current parameter file UUID: ' + MetaData['Dataset_UUID'])
    # Save to pickle file for next model run
    ParFileObject = open(ParFileName,'wb') 
    pickle.dump(ParameterDict,ParFileObject)   
    ParFileObject.close()

# ThisPar = PL_Names[mo] ThisParIx = PL_IndexStructure[mo] IndexMatch = PL_IndexMatch[mo] ThisParLayerSel = PL_IndexLayer[mo]
# Interpolate missing parameter values:
mr = 0 # reference region for GHG prices and intensities (Default: 0, which is the first region selected in the config file.)
LEDindex  = IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items.index('LED')
SSP1index = IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items.index('SSP1')
SSP2index = IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items.index('SSP2')

# 1) Material composition of vehicles, will only use historic age-cohorts.
# Values are given every 5 years, we need all values in between.
index = PL_Names.index('3_MC_RECC_Vehicles')
MC_Veh_New = np.zeros(ParameterDict[PL_Names[index]].Values.shape)
Idx_Time = [1980,1985,1990,1995,2000,2005,2010,2015,2020,2025,2030,2035,2040,2045,2050,2055,2060]
Idx_Time_Rel = [i -1900 for i in Idx_Time]
tnew = np.linspace(80, 160, num=81, endpoint=True)
for m in range(0,Nk):
    for n in range(0,Nm):
        for o in range(0,Ng):
            for p in range(0,Nr):
                f2 = interp1d(Idx_Time_Rel, ParameterDict[PL_Names[index]].Values[Idx_Time_Rel,m,n,o,p], kind='linear')
                MC_Veh_New[80::,m,n,o,p] = f2(tnew)
ParameterDict[PL_Names[index]].Values = MC_Veh_New.copy()

# 2) Material composition of buildings, will only use historic age-cohorts.
# Values are given every 5 years, we need all values in between.
index = PL_Names.index('3_MC_RECC_Buildings')
MC_Bld_New = np.zeros(ParameterDict[PL_Names[index]].Values.shape)
Idx_Time = [1900,1910,1920,1930,1940,1950,1960,1970,1980,1985,1990,1995,2000,2005,2010,2015,2020,2025,2030,2035,2040,2045,2050,2055,2060]
Idx_Time_Rel = [i -1900 for i in Idx_Time]
tnew = np.linspace(0, 160, num=161, endpoint=True)
for m in range(0,Nk):
    for n in range(0,Nm):
        for o in range(0,Ng):
            for p in range(0,Nr):
                f2 = interp1d(Idx_Time_Rel, ParameterDict[PL_Names[index]].Values[Idx_Time_Rel,m,n,o,p], kind='linear')
                MC_Bld_New[:,m,n,o,p] = f2(tnew).copy()
ParameterDict[PL_Names[index]].Values = MC_Bld_New.copy()

# 3) Determine future energy intensity and material composition of vehicles by mixing archetypes:
# Replicate values for other countries
# replicate vehicle building parameter from World to all regions
ParameterDict['3_SHA_LightWeighting_Buildings'].Values                  = np.einsum('gtS,r->grtS',ParameterDict['3_SHA_LightWeighting_Buildings'].Values[:,-1,:,:],np.ones(Nr))

# Check if RE strategies are active and set implementation curves to 2016 value if not.
if ScriptConfig['Include_REStrategy_MaterialSubstitution'] == 'False': # no lightweighting trough material substitution.
    ParameterDict['3_SHA_LightWeighting_Vehicles'].Values  = np.einsum('grS,t->grtS',ParameterDict['3_SHA_LightWeighting_Vehicles'].Values[:,:,0,:],np.ones((Nt)))
    ParameterDict['3_SHA_LightWeighting_Buildings'].Values = np.einsum('grS,t->grtS',ParameterDict['3_SHA_LightWeighting_Buildings'].Values[:,:,0,:],np.ones((Nt)))
    
if ScriptConfig['Include_REStrategy_Downsizing'] == 'False': # no lightweighting trough downsizing.
    ParameterDict['3_SHA_DownSizing_Vehicles'].Values  = np.einsum('urS,t->urtS',ParameterDict['3_SHA_DownSizing_Vehicles'].Values[:,:,0,:],np.ones((Nt)))
    ParameterDict['3_SHA_DownSizing_Buildings'].Values = np.einsum('urS,t->urtS',ParameterDict['3_SHA_DownSizing_Buildings'].Values[:,:,0,:],np.ones((Nt)))


ParameterDict['3_MC_RECC_Vehicles_RECC'] = msc.Parameter(Name='3_MC_RECC_Vehicles_RECC', ID='3_MC_RECC_Vehicles_RECC',
                                            UUID=None, P_Res=None, MetaData=None,
                                            Indices='cmgrS', Values=np.zeros((Nc,Nm,Ng,Nr,NS)), Uncert=None,
                                            Unit='kg/unit')
ParameterDict['3_MC_RECC_Vehicles_RECC'].Values[0:115,:,0:6,:,:] = np.einsum('cmgr,S->cmgrS',ParameterDict['3_MC_RECC_Vehicles'].Values[0:115,0,:,0:6,:],np.ones(NS))
ParameterDict['3_MC_RECC_Vehicles_RECC'].Values[115::,:,0:6,:,:] = \
np.einsum('grcS,gmrcS->cmgrS',ParameterDict['3_SHA_LightWeighting_Vehicles'].Values[0:6,:,:,:],np.einsum('urcS,gm->gmrcS',ParameterDict['3_SHA_DownSizing_Vehicles'].Values,ParameterDict['3_MC_VehicleArchetypes'].Values[[12,14,16,18,20,22],:]))/10000 +\
np.einsum('grcS,gmrcS->cmgrS',ParameterDict['3_SHA_LightWeighting_Vehicles'].Values[0:6,:,:,:],np.einsum('urcS,gm->gmrcS',100 - ParameterDict['3_SHA_DownSizing_Vehicles'].Values,ParameterDict['3_MC_VehicleArchetypes'].Values[[13,15,17,19,21,23],:]))/10000 +\
np.einsum('grcS,gmrcS->cmgrS',100 - ParameterDict['3_SHA_LightWeighting_Vehicles'].Values[0:6,:,:,:],np.einsum('urcS,gm->gmrcS',ParameterDict['3_SHA_DownSizing_Vehicles'].Values,ParameterDict['3_MC_VehicleArchetypes'].Values[[0,2,4,6,8,10],:]))/10000 +\
np.einsum('grcS,gmrcS->cmgrS',100 - ParameterDict['3_SHA_LightWeighting_Vehicles'].Values[0:6,:,:,:],np.einsum('urcS,gm->gmrcS',100 - ParameterDict['3_SHA_DownSizing_Vehicles'].Values,ParameterDict['3_MC_VehicleArchetypes'].Values[[1,3,5,7,9,11],:]))/10000

ParameterDict['3_EI_Products_UsePhase'].Values[115::,0:6,3,:,:,:] = \
np.einsum('grcS,gnrcS->cgnrS',ParameterDict['3_SHA_LightWeighting_Vehicles'].Values[0:6,:,:,:],np.einsum('urcS,gn->gnrcS',ParameterDict['3_SHA_DownSizing_Vehicles'].Values,ParameterDict['3_EI_VehicleArchetypes'].Values[[12,14,16,18,20,22],:]))/10000 +\
np.einsum('grcS,gnrcS->cgnrS',ParameterDict['3_SHA_LightWeighting_Vehicles'].Values[0:6,:,:,:],np.einsum('urcS,gn->gnrcS',100 - ParameterDict['3_SHA_DownSizing_Vehicles'].Values,ParameterDict['3_EI_VehicleArchetypes'].Values[[13,15,17,19,21,23],:]))/10000 +\
np.einsum('grcS,gnrcS->cgnrS',100 - ParameterDict['3_SHA_LightWeighting_Vehicles'].Values[0:6,:,:,:],np.einsum('urcS,gn->gnrcS',ParameterDict['3_SHA_DownSizing_Vehicles'].Values,ParameterDict['3_EI_VehicleArchetypes'].Values[[0,2,4,6,8,10],:]))/10000 +\
np.einsum('grcS,gnrcS->cgnrS',100 - ParameterDict['3_SHA_LightWeighting_Vehicles'].Values[0:6,:,:,:],np.einsum('urcS,gn->gnrcS',100 - ParameterDict['3_SHA_DownSizing_Vehicles'].Values,ParameterDict['3_EI_VehicleArchetypes'].Values[[1,3,5,7,9,11],:]))/10000


# 4) Determine future energy intensity and material composition of buildings by mixing archetypes:
ParameterDict['3_MC_RECC_Buildings_RECC'] = msc.Parameter(Name='3_MC_RECC_Buildings_RECC', ID='3_MC_RECC_Buildings_RECC',
                                            UUID=None, P_Res=None, MetaData=None,
                                            Indices='cmgrS', Values=np.zeros((Nc,Nm,Ng,Nr,NS)), Uncert=None,
                                            Unit='kg/unit')
ParameterDict['3_MC_RECC_Buildings_RECC'].Values[0:115,:,6::,:,:] = np.einsum('cmgr,S->cmgrS',ParameterDict['3_MC_RECC_Buildings'].Values[0:115,1,:,6::,:],np.ones(NS))
ParameterDict['3_MC_RECC_Buildings_RECC'].Values[115::,:,6::,:,:] = \
np.einsum('grcS,gmrcS->cmgrS',ParameterDict['3_SHA_LightWeighting_Buildings'].Values[6::,:,:,:],np.einsum('urcS,grm->gmrcS',ParameterDict['3_SHA_DownSizing_Buildings'].Values,ParameterDict['3_MC_BuildingArchetypes'].Values[[51,52,53,54,55,56,57,58,59],:,:]))/10000 +\
np.einsum('grcS,gmrcS->cmgrS',ParameterDict['3_SHA_LightWeighting_Buildings'].Values[6::,:,:,:],np.einsum('urcS,grm->gmrcS',100 - ParameterDict['3_SHA_DownSizing_Buildings'].Values,ParameterDict['3_MC_BuildingArchetypes'].Values[[33,34,35,36,37,38,39,40,41],:,:]))/10000 +\
np.einsum('grcS,gmrcS->cmgrS',100 - ParameterDict['3_SHA_LightWeighting_Buildings'].Values[6::,:,:,:],np.einsum('urcS,grm->gmrcS',ParameterDict['3_SHA_DownSizing_Buildings'].Values,ParameterDict['3_MC_BuildingArchetypes'].Values[[42,43,44,45,46,47,48,49,50],:,:]))/10000 +\
np.einsum('grcS,gmrcS->cmgrS',100 - ParameterDict['3_SHA_LightWeighting_Buildings'].Values[6::,:,:,:],np.einsum('urcS,grm->gmrcS',100 - ParameterDict['3_SHA_DownSizing_Buildings'].Values,ParameterDict['3_MC_BuildingArchetypes'].Values[[24,25,26,27,28,29,30,31,32],:,:]))/10000
# Replicate values for Al, Cu, Plastics:
ParameterDict['3_MC_RECC_Buildings_RECC'].Values[115::,[4,5,6,7,10],:,:,:] = np.einsum('mgr,cS->cmgrS',ParameterDict['3_MC_RECC_Buildings'].Values[110,1,[4,5,6,7,10],:,:].copy(),np.ones((Nt,NS)))
# No cement for buildings, as all cement is part of concrete:
ParameterDict['3_MC_RECC_Buildings_RECC'].Values[:,8,:,:,:] = 0

ParameterDict['3_EI_Products_UsePhase'].Values[115::,6::,:,:,:,:] = \
np.einsum('grcS,gnrVcS->cgVnrS',ParameterDict['3_SHA_LightWeighting_Buildings'].Values[6::,:,:,:],np.einsum('urcS,grVn->gnrVcS',ParameterDict['3_SHA_DownSizing_Buildings'].Values,ParameterDict['3_EI_BuildingArchetypes'].Values[[51,52,53,54,55,56,57,58,59],:,:,:]))/10000 +\
np.einsum('grcS,gnrVcS->cgVnrS',ParameterDict['3_SHA_LightWeighting_Buildings'].Values[6::,:,:,:],np.einsum('urcS,grVn->gnrVcS',100 - ParameterDict['3_SHA_DownSizing_Buildings'].Values,ParameterDict['3_EI_BuildingArchetypes'].Values[[33,34,35,36,37,38,39,40,41],:,:,:]))/10000 +\
np.einsum('grcS,gnrVcS->cgVnrS',100 - ParameterDict['3_SHA_LightWeighting_Buildings'].Values[6::,:,:,:],np.einsum('urcS,grVn->gnrVcS',ParameterDict['3_SHA_DownSizing_Buildings'].Values,ParameterDict['3_EI_BuildingArchetypes'].Values[[42,43,44,45,46,47,48,49,50],:,:,:]))/10000 +\
np.einsum('grcS,gnrVcS->cgVnrS',100 - ParameterDict['3_SHA_LightWeighting_Buildings'].Values[6::,:,:,:],np.einsum('urcS,grVn->gnrVcS',100 - ParameterDict['3_SHA_DownSizing_Buildings'].Values,ParameterDict['3_EI_BuildingArchetypes'].Values[[24,25,26,27,28,29,30,31,32],:,:,:]))/10000

# 5) Energy intensity of historic products:
#index = PL_Names.index('3_EI_Products_UsePhase_G7IC')
#ParameterDict[PL_Names[index]].Values[0:115,:,:,:,:] = np.tile(ParameterDict[PL_Names[index]].Values[115,:,:,:,:],(115,1,1,1,1))

# 6) GHG intensity of energy supply: Change unit from g/MJ to kg/MJ
index = PL_Names.index('4_PE_GHGIntensityEnergySupply')
ParameterDict[PL_Names[index]].Values = ParameterDict[PL_Names[index]].Values/1000 # convert g/MJ to kg/MJ

# 7) Fabrication yield:
# Extrapolate 2050-2060 as 2050 values
index = PL_Names.index('4_PY_Manufacturing')
ParameterDict[PL_Names[index]].Values[:,:,:,:,1::,:] = np.einsum('t,mwgFr->mwgFtr',np.ones(45),ParameterDict[PL_Names[index]].Values[:,:,:,:,0,:])

# 8) EoL RR:
ParameterDict['4_PY_EoL_RecoveryRate'].Values = np.einsum('gmwW,r->grmwW',ParameterDict['4_PY_EoL_RecoveryRate'].Values[:,-1,:,:,:],np.ones((Nr)))

# 9) Energy carrier split of buildings and vehicles
ParameterDict['3_SHA_EnergyCarrierSplit_Vehicles'].Values = np.einsum('gn,crVS->cgrVnS',ParameterDict['3_SHA_EnergyCarrierSplit_Vehicles'].Values[115,:,-1,3,:,SSP1index].copy(),np.ones((Nc,Nr,NV,NS)))

# 10) RE strategy potentials for individual countries:

ParameterDict['6_PR_ReUse_Bld'].Values                   = np.einsum('mg,r->mgr',ParameterDict['6_PR_ReUse_Bld'].Values[:,:,-1],np.ones(Nr))
ParameterDict['6_PR_LifeTimeExtension'].Values           = np.einsum('gS,r->grS',ParameterDict['6_PR_LifeTimeExtension'].Values[:,-1,:],np.ones(Nr))
ParameterDict['6_PR_FabricationYieldImprovement'].Values = np.einsum('mgS,r->mgrS',ParameterDict['6_PR_FabricationYieldImprovement'].Values[:,:,-1,:],np.ones(Nr))
ParameterDict['6_PR_EoL_RR_Improvement'].Values          = np.einsum('gmwW,r->grmwW',ParameterDict['6_PR_EoL_RR_Improvement'].Values[:,-1,:,:,:],np.ones(Nr))
ParameterDict['3_SHA_RECC_REStrategyScaleUp'].Values     = np.einsum('tR,r->trR',ParameterDict['3_SHA_RECC_REStrategyScaleUp'].Values[:,-1,:],np.ones(Nr))

# 11) LED scenario data from proxy scenarios:
# 2_P_RECC_Population_SSP_32R
ParameterDict['2_P_RECC_Population_SSP_32R'].Values[:,:,:,LEDindex]       = ParameterDict['2_P_RECC_Population_SSP_32R'].Values[:,:,:,SSP2index].copy()
# 3_EI_Products_UsePhase, historic
ParameterDict['3_EI_Products_UsePhase'].Values[0:115,:,:,:,:,LEDindex]    = ParameterDict['3_EI_Products_UsePhase'].Values[0:115,:,:,:,:,SSP2index].copy()
# 3_IO_Buildings_UsePhase_G7IC
ParameterDict['3_IO_Buildings_UsePhase'].Values[:,:,:,:,LEDindex]         = ParameterDict['3_IO_Buildings_UsePhase'].Values[:,:,:,:,SSP2index].copy()
# 4_PE_GHGIntensityEnergySupply
ParameterDict['4_PE_GHGIntensityEnergySupply'].Values[:,:,LEDindex,:,:,:] = ParameterDict['4_PE_GHGIntensityEnergySupply'].Values[:,:,SSP2index,:,:,:].copy()

#13) Combine type split parameter
ParameterDict['3_SHA_TypeSplit_NewProducts'] = msc.Parameter(Name='3_SHA_TypeSplit_NewProducts', ID='3_SHA_TypeSplit_NewProducts',
                                            UUID=None, P_Res=None, MetaData=None,
                                            Indices='GgtrS', Values=np.zeros((NG,Ng,Nt,Nr,NS)), Uncert=None,
                                            Unit='kg/unit')
ParameterDict['3_SHA_TypeSplit_NewProducts'].Values[0,0:6,:,:,:] = ParameterDict['3_SHA_TypeSplit_Vehicles'].Values[0,0:6,:,:,:].copy()
ParameterDict['3_SHA_TypeSplit_NewProducts'].Values[1,6::,:,:,:] = np.einsum('grtS->gtrS',ParameterDict['3_SHA_TypeSplit_Buildings'].Values[6::,:,:,:].copy())
# Extrapolate 2050 values for buildings
ParameterDict['3_SHA_TypeSplit_NewProducts'].Values[1,6::,36::,:,:] = np.einsum('grS,t->gtrS',ParameterDict['3_SHA_TypeSplit_NewProducts'].Values[1,6::,35,:,:].copy(),np.ones(10))

#14) Extrapolate 2050 values:
ParameterDict['2_S_RECC_FinalProducts_Future'].Values[:,36::,:,:]   = np.einsum('SGr,t->StGr',ParameterDict['2_S_RECC_FinalProducts_Future'].Values[:,35,:,:].copy(),np.ones(10))
ParameterDict['3_IO_Vehicles_UsePhase'].Values[:,:,36::,:]          = np.einsum('VrS,t->VrtS',ParameterDict['3_IO_Vehicles_UsePhase'].Values[:,:,35,:].copy(),np.ones(10))
ParameterDict['4_PE_ProcessExtensions'].Values[:,:,:,36::,:]        = np.einsum('PXrS,t->PXrtS',ParameterDict['4_PE_ProcessExtensions'].Values[:,:,:,35,:].copy(),np.ones(10)) 

#15) MODEL CALIBRATION
# Calibrate vehicle kilometrage
ParameterDict['3_IO_Vehicles_UsePhase'].Values[3,0:-1,:,:]             = ParameterDict['3_IO_Vehicles_UsePhase'].Values[3,0:-1,:,:] * np.einsum('r,tS->rtS',ParameterDict['6_PR_Calibration'].Values[0,0:-1],np.ones((Nt,NS)))
# Calibrate vehicle fuel consumption, cgVnrS
ParameterDict['3_EI_Products_UsePhase'].Values[0:115,0:6,3,:,0:-1,:]   = ParameterDict['3_EI_Products_UsePhase'].Values[0:115,0:6,3,:,0:-1,:] * np.einsum('r,cgnS->cgnrS',ParameterDict['6_PR_Calibration'].Values[1,0:-1],np.ones((115,6,Nn,NS)))
# Calibrate building energy consumption
ParameterDict['3_EI_Products_UsePhase'].Values[0:115,6::,0:3,:,0:-1,:] = ParameterDict['3_EI_Products_UsePhase'].Values[0:115,6::,0:3,:,0:-1,:] * np.einsum('r,cgVnS->cgVnrS',ParameterDict['6_PR_Calibration'].Values[2,0:-1],np.ones((115,9,3,Nn,NS)))

#16) No recycling scenario (counterfactual reference)
if ScriptConfig['IncludeRecycling'] == 'False': # no recycling and remelting
    ParameterDict['4_PY_EoL_RecoveryRate'].Values            = np.zeros(ParameterDict['4_PY_EoL_RecoveryRate'].Values.shape)
    ParameterDict['4_PY_MaterialProductionRemelting'].Values = np.zeros(ParameterDict['4_PY_MaterialProductionRemelting'].Values.shape)

# Model flow control: Include or exclude certain sectors
if ScriptConfig['SectorSelect'] == 'passenger vehicles':
    ParameterDict['2_S_RECC_FinalProducts_2015'].Values[:,:,6::,:] = 0
    ParameterDict['2_S_RECC_FinalProducts_Future'].Values[:,:,1,:] = 0
if ScriptConfig['SectorSelect'] == 'residential buildings':
    ParameterDict['2_S_RECC_FinalProducts_2015'].Values[:,:,0:6,:] = 0
    ParameterDict['2_S_RECC_FinalProducts_Future'].Values[:,:,0,:] = 0
    
Stocks_2016   = ParameterDict['2_S_RECC_FinalProducts_2015'].Values[0,:,:,:].sum(axis=0)
pCStocks_2016 = np.einsum('gr,r->rg',Stocks_2016,1/ParameterDict['2_P_RECC_Population_SSP_32R'].Values[0,0,:,1]) 
 
##########################################################
#    Section 3) Initialize dynamic MFA model for RECC    #
##########################################################
Mylog.info('## 3 - Initialize dynamic MFA model for RECC')
Mylog.info('Define RECC system and processes.')

GHG_System       = np.zeros((Nt,NS,NR))
GHG_UsePhase     = np.zeros((Nt,NS,NR))
GHG_Other        = np.zeros((Nt,NS,NR))
GHG_Materials    = np.zeros((Nt,NS,NR)) # all processes and their energy supply chains except for manufacturing and use phase
GHG_Vehicles     = np.zeros((Nt,Nr,NS,NR)) # use phase only
GHG_Buildings    = np.zeros((Nt,Nr,NS,NR)) # use phase only
GHG_Vehicles_id  = np.zeros((Nt,NS,NR)) # energy supply only
GHG_Building_id  = np.zeros((Nt,NS,NR)) # energy supply only
GHG_Manufact_all = np.zeros((Nt,NS,NR))
GHG_WasteMgt_all = np.zeros((Nt,NS,NR))
GHG_PrimaryMetal = np.zeros((Nt,NS,NR))
GHG_UsePhase_Scope2_El  = np.zeros((Nt,NS,NR))
GHG_UsePhase_OtherIndir = np.zeros((Nt,NS,NR))
GHG_MaterialCycle       = np.zeros((Nt,NS,NR))
GHG_RecyclingCredit     = np.zeros((Nt,NS,NR))
Material_Inflow  = np.zeros((Nt,Ng,Nm,NS,NR))
Scrap_Outflow    = np.zeros((Nt,Nw,NS,NR))
PrimaryProduction= np.zeros((Nt,Nm,NS,NR))
SecondaryProduct = np.zeros((Nt,Nm,NS,NR))
Element_Material_Composition     = np.zeros((Nt,Nm,Ne,NS,NR))
Element_Material_Composition_raw = np.zeros((Nt,Nm,Ne,NS,NR))
Element_Material_Composition_con = np.zeros((Nt,Nm,Ne,NS,NR))
Manufacturing_Output             = np.zeros((Nt,Ng,Nm,NS,NR))
StockMatch_2015  = np.zeros((NG,Nr))
NegInflowFlags   = np.zeros((NS,NR))
NegInflowFlags_After2020   = np.zeros((NS,NR))
FabricationScrap = np.zeros((Nt,Nw,NS,Nr))
EnergyCons_UP_Vh = np.zeros((Nt,NS,NR))
EnergyCons_UP_Bd = np.zeros((Nt,NS,NR))
EnergyCons_UP_Mn = np.zeros((Nt,NS,NR))
EnergyCons_UP_Wm = np.zeros((Nt,NS,NR))
StockCurves_Totl = np.zeros((Nt,NG,NS,NR))
StockCurves_Prod = np.zeros((Nt,Ng,NS,NR))
Population       = np.zeros((Nt,Nr,NS,NR))
pCStocksCurves   = np.zeros((Nt,NG,Nr,NS,NR))
Vehicle_km       = np.zeros((Nt,NS,NR))

#  Examples for testing
#mS = 1
#mR = 1

# Select and loop over scenarios
for mS in range(0,NS):
    for mR in range(0,NR):
        SName = IndexTable.loc['Scenario'].Classification.Items[mS]
        RName = IndexTable.loc['Scenario_RCP'].Classification.Items[mR]
        Mylog.info('Computing RECC model for SSP scenario ' + SName + ' and RE scenario ' + RName + '.')
        
        # Initialize MFA system
        RECC_System = msc.MFAsystem(Name='RECC_G7IC_SingleScenario',
                                    Geogr_Scope='14 regions', #IndexTableR.Classification[IndexTableR.set_index('IndexLetter').index.get_loc('r')].Items,
                                    Unit='Mt',
                                    ProcessList=[],
                                    FlowDict={},
                                    StockDict={},
                                    ParameterDict=ParameterDict,
                                    Time_Start=Model_Time_Start,
                                    Time_End=Model_Time_End,
                                    IndexTable=IndexTable,
                                    Elements=IndexTable.loc['Element'].Classification.Items,
                                    Graphical=None)
                              
        # Check Validity of index tables:
        # returns true if dimensions are OK and time index is present and element list is not empty
        RECC_System.IndexTableCheck()
        
        # Add processes to system
        for m in range(0, len(PrL_Number)):
            RECC_System.ProcessList.append(msc.Process(Name = PrL_Name[m], ID   = PrL_Number[m]))
            
        # Define system variables: Flows.    
        RECC_System.FlowDict['F_0_3'] = msc.Flow(Name='ore input', P_Start=0, P_End=3,
                                                 Indices='t,m,e', Values=None, Uncert=None,
                                                 Color=None, ID=None, UUID=None)
        
        RECC_System.FlowDict['F_3_4'] = msc.Flow(Name='primary material production' , P_Start = 3, P_End = 4, 
                                                 Indices = 't,m,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
        
        RECC_System.FlowDict['F_4_5'] = msc.Flow(Name='primary material consumption' , P_Start = 4, P_End = 5, 
                                                 Indices = 't,m,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
        
        RECC_System.FlowDict['F_5_6'] = msc.Flow(Name='manufacturing output' , P_Start = 5, P_End = 6, 
                                                 Indices = 't,r,g,m,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
            
        RECC_System.FlowDict['F_6_7'] = msc.Flow(Name='final consumption', P_Start=6, P_End=7,
                                                 Indices='t,r,g,m,e', Values=None, Uncert=None,
                                                 Color=None, ID=None, UUID=None)
        
        RECC_System.FlowDict['F_7_8'] = msc.Flow(Name='EoL products' , P_Start = 7, P_End = 8, 
                                                 Indices = 't,c,r,g,m,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
        
        RECC_System.FlowDict['F_8_0'] = msc.Flow(Name='obsolete stock formation' , P_Start = 8, P_End = 0, 
                                                 Indices = 't,c,r,g,m,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
        
        RECC_System.FlowDict['F_8_9'] = msc.Flow(Name='waste mgt. input' , P_Start = 8, P_End = 9, 
                                                 Indices = 't,r,g,m,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
        
        RECC_System.FlowDict['F_8_17'] = msc.Flow(Name='product re-use in' , P_Start = 8, P_End = 17, 
                                                 Indices = 't,c,r,g,m,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
        
        RECC_System.FlowDict['F_17_6'] = msc.Flow(Name='product re-use out' , P_Start = 17, P_End = 6, 
                                                 Indices = 't,c,r,g,m,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
        
        RECC_System.FlowDict['F_9_10'] = msc.Flow(Name='old scrap' , P_Start = 9, P_End = 10, 
                                                 Indices = 't,r,w,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
        
        RECC_System.FlowDict['F_5_10'] = msc.Flow(Name='new scrap' , P_Start = 5, P_End = 10, 
                                                 Indices = 't,r,w,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
        
        RECC_System.FlowDict['F_10_9'] = msc.Flow(Name='scrap use' , P_Start = 10, P_End = 9, 
                                                 Indices = 't,r,w,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
        
        RECC_System.FlowDict['F_9_12'] = msc.Flow(Name='secondary material production' , P_Start = 9, P_End = 12, 
                                                 Indices = 't,r,m,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
        
        RECC_System.FlowDict['F_12_5'] = msc.Flow(Name='secondary material consumption' , P_Start = 12, P_End = 5, 
                                                 Indices = 't,r,m,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
        
        RECC_System.FlowDict['F_12_0'] = msc.Flow(Name='excess secondary material' , P_Start = 12, P_End = 0, 
                                                 Indices = 't,r,m,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
        
        RECC_System.FlowDict['F_9_0'] = msc.Flow(Name='waste mgt. and remelting losses' , P_Start = 9, P_End = 0, 
                                                 Indices = 't,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
        
        # Define system variables: Stocks.
        RECC_System.StockDict['dS_0']  = msc.Stock(Name='System environment stock change', P_Res=0, Type=1,
                                                 Indices = 't,e', Values=None, Uncert=None,
                                                 ID=None, UUID=None)
        
        RECC_System.StockDict['S_7']   = msc.Stock(Name='In-use stock', P_Res=7, Type=0,
                                                 Indices = 't,c,r,g,m,e', Values=None, Uncert=None,
                                                 ID=None, UUID=None)
        
        RECC_System.StockDict['dS_7']  = msc.Stock(Name='In-use stock change', P_Res=7, Type=1,
                                                 Indices = 't,c,r,g,m,e', Values=None, Uncert=None,
                                                 ID=None, UUID=None)
        
        RECC_System.StockDict['S_10']   = msc.Stock(Name='Fabrication scrap buffer', P_Res=10, Type=0,
                                                 Indices = 't,c,r,w,e', Values=None, Uncert=None,
                                                 ID=None, UUID=None)
        
        RECC_System.StockDict['dS_10']  = msc.Stock(Name='Fabrication scrap buffer change', P_Res=10, Type=1,
                                                 Indices = 't,r,w,e', Values=None, Uncert=None,
                                                 ID=None, UUID=None)
        
        RECC_System.Initialize_StockValues() # Assign empty arrays to stocks according to dimensions.
        RECC_System.Initialize_FlowValues() # Assign empty arrays to flows according to dimensions.
        
        ##########################################################
        #    Section 4) Solve dynamic MFA model for RECC         #
        ##########################################################
        Mylog.info('## 4 - Solve dynamic MFA model for RECC')
        Mylog.info('Calculate inflows and outflows for use phase.')
        # THIS IS WHERE WE LEAVE THE FORMAL MODEL STRUCTURE AND DO WHATEVER IS NECESSARY TO SOLVE THE MODEL EQUATIONS.
        
        # 1) Determine total stock from regression model, and apply stock-driven model
        #TotalStockCurves_UsePhase   = np.zeros((Nt,Nr,NG))    # Stock   by year, region, and product
        #TotalStockCurves_C    = np.zeros((Nt,Nc,Nr,NG)) # Stock   by year, age-cohort, region, and product
        #TotalInflowCurves     = np.zeros((Nt,Nr,NG))    # Inflow  by year, region, and product
        #TotalOutflowCurves    = np.zeros((Nt,Nc,Nr,NG)) # Outflow by year, age-cohort, region, and product
        SF_Array               = np.zeros((Nc,Nc,Ng,Nr)) # survival functions, by year, age-cohort, good, and region. PDFs are stored externally because recreating them with scipy.stats is slow.
        Stock_Detail_UsePhase       = np.zeros((Nt,Nc,Ng,Nr)) # index structure: tcgr
        Outflow_Detail_UsePhase     = np.zeros((Nt,Nc,Ng,Nr)) # index structure: tcgr
        Inflow_Detail_UsePhase      = np.zeros((Nt,Ng,Nr)) # index structure: tgr
        SwitchTime=Nc - Model_Duration +1 # Year when future modelling horizon starts: 1.1.2016
        
        #Get historic stock at end of 2015 by age-cohort, and covert unit to Vehicles: million, Buildings: million m2.
        TotalStock_UsePhase_Hist_cgr = RECC_System.ParameterDict['2_S_RECC_FinalProducts_2015'].Values[0,:,:,:]
        
        # Determine total future stock, product level. Units: Vehicles: million, Buildings: million m2.
        TotalStockCurves_UsePhase = np.einsum('tGr,tr->trG',RECC_System.ParameterDict['2_S_RECC_FinalProducts_Future'].Values[mS,:,:,:],RECC_System.ParameterDict['2_P_RECC_Population_SSP_32R'].Values[0,:,:,mS]) 
        # Here, the population model M is set to its default and does not appear in the summation.
        
        # 2) Include (or not) the RE strategies for the use phase:
        # Include_REStrategy_MoreIntenseUse for SSP1 and SSP2:
        if ScriptConfig['Include_REStrategy_MoreIntenseUse'] == 'True': # calculate counter-factual scenario: SSP1 stocks for SSP2, LED stocks for SSP1
            if IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items[mS] == 'SSP2': # use SSP1 stock curves instead:
                MS_SSP1 = IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items.index('SSP1')
                TotalStockCurves_UsePhase = np.einsum('tGr,tr->trG',RECC_System.ParameterDict['2_S_RECC_FinalProducts_Future'].Values[MS_SSP1,:,:,:],RECC_System.ParameterDict['2_P_RECC_Population_SSP_32R'].Values[0,:,:,mS])         
            if IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items[mS] == 'SSP1': # use LED stock curves instead:
                MS_LED = IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items.index('LED')
                TotalStockCurves_UsePhase = np.einsum('tGr,tr->trG',RECC_System.ParameterDict['2_S_RECC_FinalProducts_Future'].Values[MS_LED,:,:,:],RECC_System.ParameterDict['2_P_RECC_Population_SSP_32R'].Values[0,:,:,mS])         

        # Include_REStrategy_LifeTimeExtension: Product lifetime extension.
        # First, replicate lifetimes for all age-cohorts
        Par_RECC_ProductLifetime = np.einsum('c,gUr->grc',np.ones((Nc)),RECC_System.ParameterDict['3_LT_RECC_ProductLifetime'].Values) # Sums up over U, only possible because of 1:1 correspondence of U and g!
        # Second, change lifetime of future age-cohorts according to lifetime extension parameter
        # This is equation 10 of the paper:
        if ScriptConfig['Include_REStrategy_LifeTimeExtension'] == 'True':
            Par_RECC_ProductLifetime[:,:,SwitchTime -1::] = np.einsum('crg,grc->grc',1 + np.einsum('cr,gr->crg',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp'].Values[:,:,mR]*0.01,RECC_System.ParameterDict['6_PR_LifeTimeExtension'].Values[:,:,mS]),Par_RECC_ProductLifetime[:,:,SwitchTime -1::])
        
        # 3) Dynamic stock model
        # Build pdf array from lifetime distribution: Probability of survival.
        for g in tqdm(range(0, Ng), unit=' commodity groups'):
            for r in range(0, Nr):
                LifeTimes = Par_RECC_ProductLifetime[g, r, :]
                lt = {'Type'  : 'Normal',
                      'Mean'  : LifeTimes,
                      'StdDev': 0.3 * LifeTimes}
                SF_Array[:, :, g, r] = dsm.DynamicStockModel(t=np.arange(0, Nc, 1), lt=lt).compute_sf().copy()
                np.fill_diagonal(SF_Array[:, :, g, r],1) # no outflows from current year, this would break the mass balance in the calculation routine below, as the element composition of the current year is not yet known.
                # Those parts of the stock remain in use instead.

        # Compute evolution of 2015 in-use stocks: initial stock evolution separately from future stock demand and stock-driven model
        for G in tqdm(range(0, NG), unit=' commodity groups'):
            for r in range(0,Nr):   
                FutureStock          = TotalStockCurves_UsePhase[1::, r, G]# Future total stock
                InitialStock         = TotalStock_UsePhase_Hist_cgr[:,:,r].copy()
                if G == 0: # for vehicles and buildings only !!!
                    InitialStock[:,6::] = 0  # set not relevant initial stock to 0
                    InitialStocksum     = InitialStock[:,0:6].sum()
                if G == 1:
                    InitialStock[:,0:6] = 0  # set not relevant initial stock to 0
                    InitialStocksum     = InitialStock[:,6::].sum()
                StockMatch_2015[G,r] = TotalStockCurves_UsePhase[0, r, G]/InitialStocksum
                SFArrayCombined = SF_Array[:,:,:,r]
                TypeSplit       = RECC_System.ParameterDict['3_SHA_TypeSplit_NewProducts'].Values[G,:,1::,r,mS].transpose() # indices: gc
  
                Var_S, Var_O, Var_I = msf.compute_stock_driven_model_initialstock_typesplit(FutureStock,InitialStock,SFArrayCombined,TypeSplit, NegativeInflowCorrect = True)

                # Below, the results are added with += because the different commodity groups (buildings, vehicles) are calculated separately
                # to introduce the type split for each, but using the product resolution of the full model with all sectors.
                Stock_Detail_UsePhase[0,:,:,r]     += InitialStock.copy() # cgr, needed for correct calculation of mass balance later.
                Stock_Detail_UsePhase[1::,:,:,r]   += Var_S.copy() # tcgr
                Outflow_Detail_UsePhase[1::,:,:,r] += Var_O.copy() # tcgr
                Inflow_Detail_UsePhase[1::,:,r]    += Var_I[SwitchTime::,:].copy() # tgr

        # Here so far: Units: Vehicles: million, Buildings: million m2. for stocks, X/yr for flows.
        StockCurves_Totl[:,:,mS,mR] = TotalStockCurves_UsePhase.sum(axis =1)
        StockCurves_Prod[:,:,mS,mR] = np.einsum('tcgr->tg',Stock_Detail_UsePhase)
        pCStocksCurves[:,:,:,mS,mR] = ParameterDict['2_S_RECC_FinalProducts_Future'].Values[mS,:,:,:]
        Population[:,:,mS,mR]       = ParameterDict['2_P_RECC_Population_SSP_32R'].Values[0,:,:,mS]
        
        # Clean up
        #del TotalStockCurves_UsePhase
        #del SF_Array
        
        # Flag scenario with negative inflows:
        if (Inflow_Detail_UsePhase < 0).sum() > 0:
            NegInflowFlags[mS,mR] = 1
        if (Inflow_Detail_UsePhase[5::,:,:] < 0).sum() > 0:
            NegInflowFlags_After2020[mS,mR] = 1
        
        # Prepare parameters:        
        # include light-weighting in future MC parameter, cmgr
        Par_RECC_MC = RECC_System.ParameterDict['3_MC_RECC_Vehicles_RECC'].Values[:,:,:,:,mS] + RECC_System.ParameterDict['3_MC_RECC_Buildings_RECC'].Values[:,:,:,:,mS]
        
        # Units: Vehicles: kg/unit, Buildings: kg/m2  
        
        # historic element composition of materials:
        Par_Element_Composition_of_Materials_m   = np.zeros((Nc,Nm,Ne)) # cme, produced in age-cohort c. Applies to new manufactured goods.
        Par_Element_Composition_of_Materials_m[0:Nc-Nt+1,:,:] = np.einsum('c,me->cme',np.ones(Nc-Nt+1),RECC_System.ParameterDict['3_MC_Elements_Materials_ExistingStock'].Values)
        # For future age-cohorts, the total is known but the element breakdown of this parameter will be updated year by year in the loop below.
        Par_Element_Composition_of_Materials_m[:,:,0] = 1 # element 0 is 'all', for which the mass share is 100%.
        
        # future element composition of materials inflow use phase (mix new and reused products)
        Par_Element_Composition_of_Materials_c   = np.zeros((Nt,Nm,Ne)) # cme, produced in age-cohort c. Applies to new manufactured goods.
        
        # Element composition of material in the use phase
        Par_Element_Composition_of_Materials_u   = Par_Element_Composition_of_Materials_m.copy() # cme
        
        # Manufacturing yield and other improvements
        Par_FabYield = np.einsum('mwggtr->mwgtr',RECC_System.ParameterDict['4_PY_Manufacturing'].Values) # take diagonal of product = manufacturing process
        # Consider Fabrication yield improvement
        if ScriptConfig['Include_REStrategy_FabYieldImprovement'] == 'True':
            Par_FabYieldImprovement = np.einsum('w,tmgr->mwgtr',np.ones((Nw)),np.einsum('tr,mgr->tmgr',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp'].Values[:,:,mR]*0.01,RECC_System.ParameterDict['6_PR_FabricationYieldImprovement'].Values[:,:,:,mS]))
            # Reduce cement content by up to 15.6 %
            Par_RECC_MC[115::,[11,12],:,:] = Par_RECC_MC[115::,[11,12],:,:] * (1 - 0.156 * np.einsum('tr,mg->tmgr',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp'].Values[:,:,mR]*0.01,np.ones((2,Ng)))).copy()
        else:
            Par_FabYieldImprovement = 0

        #Mylog.info('Concrete Content total: %s' % (Par_RECC_MC[130::,11,:,:].sum()))                    
            
        Par_FabYield_Raster = Par_FabYield > 0    
        Par_FabYield        = Par_FabYield - Par_FabYield_Raster * Par_FabYieldImprovement #mwgtr
        Par_FabYield_total  = np.einsum('mwgtr->mgtr',Par_FabYield)
        Par_FabYield_total_inv = 1/(1-Par_FabYield_total) # mgtr
        
        # Determine total element composition of products, needs to be updated for future age-cohorts!
        Par_Element_Material_Composition_of_Products = np.einsum('cmgr,cme->crgme',Par_RECC_MC,Par_Element_Composition_of_Materials_m)
        
        # Consider EoL recovery rate improvement:
        if ScriptConfig['Include_REStrategy_EoL_RR_Improvement'] == 'True':
            Par_RECC_EoL_RR =  np.einsum('t,grmw->trmgw',np.ones((Nt)),RECC_System.ParameterDict['4_PY_EoL_RecoveryRate'].Values[:,:,:,:,0] *0.01) \
            + np.einsum('tr,grmw->trmgw',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp'].Values[:,:,mR]*0.01,RECC_System.ParameterDict['6_PR_EoL_RR_Improvement'].Values[:,:,:,:,0]*0.01)
        else:    
            Par_RECC_EoL_RR =  np.einsum('t,grmw->trmgw',np.ones((Nt)),RECC_System.ParameterDict['4_PY_EoL_RecoveryRate'].Values[:,:,:,:,0] *0.01)
        
        # Calculate reuse factor
        # For vehicles
        ReUseFactor_tmgrS = np.einsum('mgrtS->tmgrS',RECC_System.ParameterDict['6_PR_ReUse_Veh'].Values/100)
        # For Buildings
        ReUseFactor_tmgrS[:,:,6::,:,:] = np.einsum('tmgr,S->tmgrS',np.einsum('tr,mgr->tmgr',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp'].Values[:,:,mR]*0.01,RECC_System.ParameterDict['6_PR_ReUse_Bld'].Values[:,6::,:]),np.ones((NS)))
        
        Mylog.info('Translate total flows into individual materials and elements, for 2015 and historic age-cohorts.')
        
        # 1) Inflow, outflow, and stock first year
        RECC_System.FlowDict['F_6_7'].Values[0,:,:,:,:]   = \
        np.einsum('rgme,gr->rgme',Par_Element_Material_Composition_of_Products[SwitchTime-1,:,:,:,:],Inflow_Detail_UsePhase[0,:,:])/1000 # all elements, Indices='t,r,g,m,e'  
        
        RECC_System.FlowDict['F_7_8'].Values[0,0:SwitchTime,:,:,:,:] = \
        np.einsum('crgme,cgr->crgme',Par_Element_Material_Composition_of_Products[0:SwitchTime,:,:,:,:],Outflow_Detail_UsePhase[0,0:SwitchTime,:,:])/1000 # all elements, Indices='t,r,g,m,e'
        
        RECC_System.StockDict['S_7'].Values[0,0:SwitchTime,:,:,:,:] = \
        np.einsum('crgme,cgr->crgme',Par_Element_Material_Composition_of_Products[0:SwitchTime,:,:,:,:],Stock_Detail_UsePhase[0,0:SwitchTime,:,:])/1000 # all elements, Indices='t,r,g,m,e'

        # 1) Inflow, future years, all elements only
        RECC_System.FlowDict['F_6_7'].Values[1::,:,:,:,0]   = \
        np.einsum('trgm,tgr->trgm',Par_Element_Material_Composition_of_Products[SwitchTime::,:,:,:,0],Inflow_Detail_UsePhase[1::,:,:])/1000 # all elements, Indices='t,r,g,m,e'  
                
        #Units so far: Mt/yr
        
        Mylog.info('Calculate material stocks and flows, material cycles, determine elemental composition.')
        # Units: Mt and Mt/yr.
        # This calculation is done year-by-year, and the elemental composition of the materials is in part determined by the scrap flow metal composition
        
        for t in tqdm(range(1, Nt), unit=' years'): # 1: 2016
        #for t in tqdm(range(1, 5), unit=' years'): # 1: 2016
            CohortOffset = t +Nc -Nt # index of current age-cohort.   
            # First, before going down to the material layer, we consider obsolete stock formation and re-use.
            
            # 1) Convert use phase outflow to system variables.
            # Split flows into materials and chemical elements.
            # Calculate use phase outflow and obsolete stock formation
            # ObsStockFormation = ObsStockFormationFactor(t,g,r) * Outflow_Detail_UsePhase(t,c,g,r), currently not implemented. 
            RECC_System.FlowDict['F_7_8'].Values[t,0:CohortOffset,:,:,:,:] = \
            np.einsum('crgme,cgr->crgme',Par_Element_Material_Composition_of_Products[0:CohortOffset,:,:,:,:],Outflow_Detail_UsePhase[t,0:CohortOffset,:,:])/1000 # All elements.
            # RECC_System.FlowDict['F_8_0'].Values = MatContent * ObsStockFormation. Currently 0, already defined.
                        
            # 2) Consider re-use of materials in product groups (via components), as ReUseFactor(m,g,r,R,t) * RECC_System.FlowDict['F_7_8'].Values(t,c,r,g,m,e)
            # Distribute material for re-use onto product groups
            if ScriptConfig['Include_REStrategy_ReUse'] == 'True':
                ReUsePotential_Materials_t_m_Veh = np.einsum('mgr,crgm->m',ReUseFactor_tmgrS[t,:,0:6,:,mS],RECC_System.FlowDict['F_7_8'].Values[t,:,:,0:6,:,0])
                ReUsePotential_Materials_t_m_Bld = np.einsum('mg,crgm->m',ReUseFactor_tmgrS[t,:,6::,-1,mS],RECC_System.FlowDict['F_7_8'].Values[t,:,:,6::,:,0])
                # in the future, re-use will be a region-to-region parameter depicting, e.g., the export of used vehicles from the EU to Africa.
                # check whether inflow is big enough for potential to be used, correct otherwise:
                for mmm in range(0,Nm):
                    # Vehicles
                    if RECC_System.FlowDict['F_6_7'].Values[t,:,0:6,mmm,0].sum() < ReUsePotential_Materials_t_m_Veh[mmm]: # if re-use potential is larger than new inflow:
                        if ReUsePotential_Materials_t_m_Veh[mmm] > 0:
                            ReUsePotential_Materials_t_m_Veh[mmm] = RECC_System.FlowDict['F_6_7'].Values[t,:,0:6,mmm,0].sum()
                    # Buildings
                    if RECC_System.FlowDict['F_6_7'].Values[t,:,6::,mmm,0].sum() < ReUsePotential_Materials_t_m_Bld[mmm]: # if re-use potential is larger than new inflow:
                        if ReUsePotential_Materials_t_m_Bld[mmm] > 0:
                            ReUsePotential_Materials_t_m_Bld[mmm] = RECC_System.FlowDict['F_6_7'].Values[t,:,6::,mmm,0].sum()
            else:
                ReUsePotential_Materials_t_m_Veh = np.zeros((Nm)) # in Mt/yr, total mass
                ReUsePotential_Materials_t_m_Bld = np.zeros((Nm)) # in Mt/yr, total mass
            
            # Vehicles
            MassShareVeh = RECC_System.FlowDict['F_7_8'].Values[t,0:CohortOffset,:,0:6,:,0] / np.einsum('m,crg->crgm',np.einsum('crgm->m',RECC_System.FlowDict['F_7_8'].Values[t,0:CohortOffset,:,0:6,:,0]),np.ones((CohortOffset,Nr,6)))
            MassShareVeh[np.isnan(MassShareVeh)] = 0 # share of combination crg in total mass of m in outflow 7_8
            RECC_System.FlowDict['F_8_17'].Values[t,0:CohortOffset,:,0:6,:,:] = \
            np.einsum('cme,crgm->crgme', Par_Element_Composition_of_Materials_u[0:CohortOffset,:,:],\
            np.einsum('m,crgm->crgm',ReUsePotential_Materials_t_m_Veh,MassShareVeh))  # All elements.
            # Buildings
            MassShareBld = RECC_System.FlowDict['F_7_8'].Values[t,0:CohortOffset,:,6::,:,0] / np.einsum('m,crg->crgm',np.einsum('crgm->m',RECC_System.FlowDict['F_7_8'].Values[t,0:CohortOffset,:,6::,:,0]),np.ones((CohortOffset,Nr,9)))
            MassShareBld[np.isnan(MassShareBld)] = 0 # share of combination crg in total mass of m in outflow 7_8
            RECC_System.FlowDict['F_8_17'].Values[t,0:CohortOffset,:,6::,:,:] = \
            np.einsum('cme,crgm->crgme', Par_Element_Composition_of_Materials_u[0:CohortOffset,:,:],\
            np.einsum('m,crgm->crgm',ReUsePotential_Materials_t_m_Bld,MassShareBld))  # All elements.
            
            InvMass = 1 / np.einsum('m,rg->rgm',np.einsum('rgm->m',RECC_System.FlowDict['F_6_7'].Values[t,:,:,:,0]),np.ones((Nr,Ng)))
            InvMass[np.isnan(InvMass)] = 0
            InvMass[np.isinf(InvMass)] = 0
            RECC_System.FlowDict['F_17_6'].Values[t,0:CohortOffset,:,:,:,:] = \
            np.einsum('cme,rgm->crgme',np.einsum('crgme->cme',RECC_System.FlowDict['F_8_17'].Values[t,0:CohortOffset,:,:,:,:]),\
            RECC_System.FlowDict['F_6_7'].Values[t,:,:,:,0]*InvMass) # reused material mapped to final consumption region and good
            
            # 3) Add re-use flow to inflow and calculate manufacturing output, in Mt/yr, all elements, trgme, element composition not yet known.
            Manufacturing_Output_gm           = np.einsum('rgm->gm',RECC_System.FlowDict['F_6_7'].Values[t,:,:,:,0]) - np.einsum('crgm->gm',RECC_System.FlowDict['F_17_6'].Values[t,:,:,:,:,0]) # global total
            Manufacturing_Output[t,:,:,mS,mR] = Manufacturing_Output_gm
            
            # 4) calculate inflow waste mgt.
            RECC_System.FlowDict['F_8_9'].Values[t,:,:,:,:]     = np.einsum('crgme->rgme',RECC_System.FlowDict['F_7_8'].Values[t,0:CohortOffset,:,:,:,:] - RECC_System.FlowDict['F_8_0'].Values[t,0:CohortOffset,:,:,:,:] - RECC_System.FlowDict['F_8_17'].Values[t,0:CohortOffset,:,:,:,:])
    
            # 5) EoL products to postconsumer scrap: trwe
            PostConsumerScrap_ByRegion                          = np.einsum('rmgw,rgme->rwe',Par_RECC_EoL_RR[t,:,:,:,:],RECC_System.FlowDict['F_8_9'].Values[t,:,:,:,:])    
            # Aggregate scrap flows at world level:
            RECC_System.FlowDict['F_9_10'].Values[t,-1,:,:]     = np.einsum('rme->me',PostConsumerScrap_ByRegion)
               
            # 6) Add new scrap and calculate remelting.
            # Add old scrap with manufacturing scrap from last year. In year 2016, no fabrication scrap exists yet.
            RECC_System.FlowDict['F_10_9'].Values[t,:,:,:]      = np.einsum('rwe->rwe',RECC_System.FlowDict['F_9_10'].Values[t,:,:,:] + RECC_System.StockDict['S_10'].Values[t-1,t-1,:,:,:].copy())
            RECC_System.FlowDict['F_9_12'].Values[t,:,:,:]      = np.einsum('rwe,wmePr->rme',RECC_System.FlowDict['F_10_9'].Values[t,:,:,:],RECC_System.ParameterDict['4_PY_MaterialProductionRemelting'].Values[:,:,:,:,0,:])
            RECC_System.FlowDict['F_9_12'].Values[t,:,:,0]      = np.einsum('rme->rm',RECC_System.FlowDict['F_9_12'].Values[t,:,:,1::])
            RECC_System.FlowDict['F_12_5'].Values[t,:,:,:]      = RECC_System.FlowDict['F_9_12'].Values[t,:,:,:]
            RECC_System.FlowDict['F_12_5'].Values[t,:,:,0]      = np.einsum('rme->rm',RECC_System.FlowDict['F_12_5'].Values[t,:,:,1::]) # All up all chemical elements to total
                     
            # Element composition shares of recycled material:
            Element_Material_Composition_t_SecondaryMaterial = np.einsum('me,me->me',RECC_System.FlowDict['F_9_12'].Values[t,-1,:,:],1/np.einsum('m,e->me',RECC_System.FlowDict['F_9_12'].Values[t,-1,:,0],np.ones(Ne)))
            Element_Material_Composition_t_SecondaryMaterial[np.isnan(Element_Material_Composition_t_SecondaryMaterial)] = 0            
            
            # 7) Waste mgt. losses.
            RECC_System.FlowDict['F_9_0'].Values[t,:]         = np.einsum('rgme->e',RECC_System.FlowDict['F_8_9'].Values[t,:,:,:,:]) + np.einsum('rwe->e',RECC_System.FlowDict['F_10_9'].Values[t,:,:,:]) - np.einsum('rwe->e',RECC_System.FlowDict['F_9_10'].Values[t,:,:,:]) - np.einsum('rme->e',RECC_System.FlowDict['F_9_12'].Values[t,:,:,:])

            # 8) Calculate manufacturing input and primary production, all elements, element composition not yet known.
            Manufacturing_Input_m        = np.einsum('mg,gm->m',Par_FabYield_total_inv[:,:,t,-1],Manufacturing_Output_gm)
            Manufacturing_Input_gm       = np.einsum('mg,gm->gm',Par_FabYield_total_inv[:,:,t,-1],Manufacturing_Output_gm)
            Manufacturing_Input_Split_gm = np.einsum('gm,m->gm',Manufacturing_Input_gm, 1/Manufacturing_Input_m)
            Manufacturing_Input_Split_gm[np.isnan(Manufacturing_Input_Split_gm)] = 0
            
            PrimaryProduction_m          = Manufacturing_Input_m - RECC_System.FlowDict['F_12_5'].Values[t,-1,:,0]# secondary material comes first, no rebound! 
            
            # Correct for negative primary production: p.p. is set to zero, and a corresponding quantity is exported instead:
            for pm in range(0,Nm):
                if PrimaryProduction_m[pm] < 0:
                    RECC_System.FlowDict['F_12_0'].Values[t,-1,pm,0] = -1 * PrimaryProduction_m[pm].copy()
                    RECC_System.FlowDict['F_12_0'].Values[t,-1,pm,:] = Element_Material_Composition_t_SecondaryMaterial[pm,:] * RECC_System.FlowDict['F_12_0'].Values[t,-1,pm,0]
                    RECC_System.FlowDict['F_12_5'].Values[t,-1,pm,:] = RECC_System.FlowDict['F_12_5'].Values[t,-1,pm,:] - RECC_System.FlowDict['F_12_0'].Values[t,-1,pm,:]
                    PrimaryProduction_m[pm] = 0
        
            RECC_System.FlowDict['F_4_5'].Values[t,:,:] = np.einsum('m,me->me',PrimaryProduction_m,RECC_System.ParameterDict['3_MC_Elements_Materials_Primary'].Values)
            RECC_System.FlowDict['F_3_4'].Values[t,:,:] = RECC_System.FlowDict['F_4_5'].Values[t,:,:]
            RECC_System.FlowDict['F_0_3'].Values[t,:,:] = RECC_System.FlowDict['F_3_4'].Values[t,:,:]
        
            Manufacturing_Input_me      = RECC_System.FlowDict['F_4_5'].Values[t,:,:] + RECC_System.FlowDict['F_12_5'].Values[t,-1,:,:] 
            Manufacturing_Input_gme     = np.einsum('me,gm->gme',Manufacturing_Input_me,Manufacturing_Input_Split_gm)       
        
            # 9) Calculate element composition of materials of current year
            Element_Material_Composition_t = np.einsum('me,me->me',RECC_System.FlowDict['F_4_5'].Values[t,:,:] + RECC_System.FlowDict['F_12_5'].Values[t,-1,:,:],1/np.einsum('m,e->me',RECC_System.FlowDict['F_4_5'].Values[t,:,0] + RECC_System.FlowDict['F_12_5'].Values[t,-1,:,0],np.ones(Ne)))
            Element_Material_Composition_raw[t,:,:,mS,mR] = Element_Material_Composition_t.copy()
            Element_Material_Composition_t[np.isnan(Element_Material_Composition_t)] = 0
            Element_Material_Composition_t[np.isinf(Element_Material_Composition_t)] = 0
            #Element_Material_Composition_t[:,0] = 1
            #Element_Material_Composition_t[:,-1] = 1 - Element_Material_Composition_t[:,1:-1].sum(axis =1)
            Element_Material_Composition[t,:,:,mS,mR] = Element_Material_Composition_t.copy()
            Par_Element_Composition_of_Materials_m[t+115,:,:] = Element_Material_Composition_t.copy()
            Par_Element_Composition_of_Materials_u[t+115,:,:] = Element_Material_Composition_t.copy()
            
            # 10) Calculate manufacturing output, at global level only
            RECC_System.FlowDict['F_5_6'].Values[t,-1,:,:,:] = np.einsum('me,gm->gme',Element_Material_Composition_t,Manufacturing_Output_gm)
        
            # Manufacturing diagnostics
            Aa = RECC_System.FlowDict['F_5_6'].Values[:,-1,:,:,:].sum(axis=1) #tme
            Ab = RECC_System.FlowDict['F_5_10'].Values[:,-1,:,:] # twe
            Ac = RECC_System.FlowDict['F_4_5'].Values # tme
            Ad = RECC_System.FlowDict['F_12_5'].Values[:,-1,:,:] # tme
            Bal_20 = Aa[20,:,0] - Ac[20,:,0] - Ad[20,:,0]
            Bal_30 = Aa[30,:,0] - Ac[30,:,0] - Ad[30,:,0]
        
            # 10a) Calculate material composition of product consumption
            Throughput_FinalGoods_me = RECC_System.FlowDict['F_5_6'].Values[t,-1,:,:,:].sum(axis =0) + np.einsum('crgme->me',RECC_System.FlowDict['F_17_6'].Values[t,0:CohortOffset,:,:,:,:])
            Element_Material_Composition_cons = np.einsum('me,me->me',Throughput_FinalGoods_me,1/np.einsum('m,e->me',Throughput_FinalGoods_me[:,1::].sum(axis =1),np.ones(Ne)))
            Element_Material_Composition_cons[np.isnan(Element_Material_Composition_cons)] = 0
            Element_Material_Composition_cons[np.isinf(Element_Material_Composition_cons)] = 0
            #Element_Material_Composition_cons[:,0] = 1
            #Element_Material_Composition_cons[:,-1] = 1 - Element_Material_Composition_t[:,1:-1].sum(axis =1)
            Element_Material_Composition_con[t,:,:,mS,mR] = Element_Material_Composition_cons.copy()
            
            Par_Element_Composition_of_Materials_c[t,:,:] = Element_Material_Composition_cons.copy()
            Par_Element_Material_Composition_of_Products[CohortOffset,:,:,:,:] = np.einsum('mgr,me->rgme',Par_RECC_MC[CohortOffset,:,:,:],Par_Element_Composition_of_Materials_c[t,:,:]) # crgme
                    
            # 11) Calculate manufacturing scrap 
            RECC_System.FlowDict['F_5_10'].Values[t,-1,:,:] = np.einsum('gme,mwgr->we',Manufacturing_Input_gme,Par_FabYield[:,:,:,t,:]) 
            # Fabrication scrap, to be recycled next year:
            RECC_System.StockDict['S_10'].Values[t,t,:,:,:]  = RECC_System.FlowDict['F_5_10'].Values[t,:,:,:]
        
            # 12) Calculate element composition of final consumption and latest age-cohort in in-use stock
            RECC_System.FlowDict['F_6_7'].Values[t,:,:,:,:]   = \
            np.einsum('rgme,gr->rgme',Par_Element_Material_Composition_of_Products[CohortOffset,:,:,:,:],Inflow_Detail_UsePhase[t,:,:])/1000 # all elements, Indices='t,r,g,m,e'
            
            RECC_System.StockDict['S_7'].Values[t,0:CohortOffset +1,:,:,:,:] = \
            np.einsum('crgme,cgr->crgme',Par_Element_Material_Composition_of_Products[0:CohortOffset +1,:,:,:,:],Stock_Detail_UsePhase[t,0:CohortOffset +1,:,:])/1000 # All elements.
 
            # 13) Calculate stock changes
            RECC_System.StockDict['dS_7'].Values[t,:,:,:,:,:] = RECC_System.StockDict['S_7'].Values[t,:,:,:,:,:] - RECC_System.StockDict['S_7'].Values[t-1,:,:,:,:,:]
            RECC_System.StockDict['dS_10'].Values[t,:,:,:]    = RECC_System.StockDict['S_10'].Values[t,t,:,:,:] - RECC_System.StockDict['S_10'].Values[t-1,t-1,:,:,:]
            RECC_System.StockDict['dS_0'].Values[t,:]         = RECC_System.FlowDict['F_9_0'].Values[t,:] + np.einsum('rme->e',RECC_System.FlowDict['F_12_0'].Values[t,:,:,:]) + np.einsum('crgme->e',RECC_System.FlowDict['F_8_0'].Values[t,:,:,:,:,:]) - np.einsum('me->e',RECC_System.FlowDict['F_0_3'].Values[t,:,:])
            
            
            # Diagnostics:
            Aa = np.einsum('tcrgm->trm',RECC_System.FlowDict['F_7_8'].Values[:,:,:,6::,:,0]) # BuildingOutflowMaterials
            Aa = np.einsum('tcrgm->trm',RECC_System.FlowDict['F_7_8'].Values[:,:,:,0:6,:,0]) # VehiclesOutflowMaterials
            Aa = np.einsum('tcrgm->tmr',RECC_System.FlowDict['F_7_8'].Values[:,:,:,:,:,0]) # outflow use phase
            Aa = np.einsum('tcrgm->tr',RECC_System.FlowDict['F_8_17'].Values[:,:,:,:,:,0]) # reuse
            Aa = np.einsum('trw->tw',RECC_System.FlowDict['F_9_10'].Values[:,:,:,0])       # old scrap
            Aa = np.einsum('trgm->tr',RECC_System.FlowDict['F_8_9'].Values[:,:,:,:,0])     # inflow waste mgt.
            Aa = np.einsum('trm->tm',RECC_System.FlowDict['F_9_12'].Values[:,:,:,0])       # secondary material
            Aa = np.einsum('tcgr->tgr',Outflow_Detail_UsePhase)                            # product outflow use phase
            Aa = np.einsum('tgr->tgr',Inflow_Detail_UsePhase)                              # product inflow use phase
            Aa = np.einsum('tgr->tg',Inflow_Detail_UsePhase)                               # product inflow use phase, global total
            Aa = np.einsum('trm->trm',RECC_System.FlowDict['F_12_5'].Values[:,:,:,0])      # secondary material use
            Aa = np.einsum('tm->tm',RECC_System.FlowDict['F_4_5'].Values[:,:,0])           # primary material production
            Aa = np.einsum('trgm->trm',RECC_System.FlowDict['F_5_6'].Values[:,:,0:6,:,0])  # materials in manufactured vehicles
            Aa = np.einsum('trgm->trm',RECC_System.FlowDict['F_5_6'].Values[:,:,6::,:,0])  # materials in manufactured buildings
            Aa = np.einsum('tcgr->tgr',Stock_Detail_UsePhase)                              # Total stock time series
            
            # ReUse Diagnostics
            Ab = np.einsum('tcrgme->tme',RECC_System.FlowDict['F_8_17'].Values)
            Ab = np.einsum('tcrgme->tme',RECC_System.FlowDict['F_17_6'].Values)
            Ab = Element_Material_Composition_con[:,:,:,mS,mR]
            Ab = MassShareVeh[:,0,:,:]
            Ab = np.einsum('m,crgm->cgm',ReUsePotential_Materials_t_m_Veh,MassShareVeh)
            
        # Check whether flow value arrays match their indices, etc.
        RECC_System.Consistency_Check() 
    
        # Determine Mass Balance
        Bal = RECC_System.MassBalance()
        BalAbs = np.abs(Bal).sum()
        Mylog.info('Total mass balance deviation (np.abs(Bal).sum() for socioeconomic scenario ' + SName + ' and RE scenario ' + RName + ': ' + str(BalAbs) + ' Mt.')                    
        
        # A) Calculate intensity of operation
        SysVar_StockServiceProvision_UsePhase = np.einsum('Vrt,tcgr->tcgrV',RECC_System.ParameterDict['3_IO_Vehicles_UsePhase'].Values[:,:,:,mS], Stock_Detail_UsePhase) + np.einsum('cgVr,tcgr->tcgrV',RECC_System.ParameterDict['3_IO_Buildings_UsePhase'].Values[:,:,:,:,mS], Stock_Detail_UsePhase)
        # Unit: million km/yr for vehicles, million m2 for buildings by three use types: heating, cooling, and DHW.
        
        # B) Calculate total operational energy use
        SysVar_EnergyDemand_UsePhase_Total  = np.einsum('cgVnr,tcgrV->tcgrnV',RECC_System.ParameterDict['3_EI_Products_UsePhase'].Values[:,:,:,:,:,mS], SysVar_StockServiceProvision_UsePhase)
        # Unit: TJ/yr for both vehicles and buildings.
        
        # C) Translate 'all' energy carriers to specific ones, use phase
        SysVar_EnergyDemand_UsePhase_Buildings_ByEnergyCarrier = np.einsum('Vrnt,tcgrV->trgn',RECC_System.ParameterDict['3_SHA_EnergyCarrierSplit_Buildings'].Values[:,mR,:,:,:],SysVar_EnergyDemand_UsePhase_Total[:,:,:,:,-1,:].copy())
        SysVar_EnergyDemand_UsePhase_Vehicles_ByEnergyCarrier  = np.einsum('cgrVn,tcgrV->trgn',RECC_System.ParameterDict['3_SHA_EnergyCarrierSplit_Vehicles'].Values[:,:,:,:,:,mS] ,SysVar_EnergyDemand_UsePhase_Total[:,:,:,:,-1,:].copy())
        
        # D) Calculate energy demand of the other industries
        SysVar_EnergyDemand_PrimaryProd   = 1000 * np.einsum('mn,tm->tmn',RECC_System.ParameterDict['4_EI_ProcessEnergyIntensity'].Values[:,:,110,0],RECC_System.FlowDict['F_3_4'].Values[:,:,0])
        SysVar_EnergyDemand_Manufacturing = 1 * np.einsum('gn,tgr->tn',RECC_System.ParameterDict['4_EI_ManufacturingEnergyIntensity'].Values[:,:,110,-1],Inflow_Detail_UsePhase)
        SysVar_EnergyDemand_WasteMgt      = 1000 * np.einsum('wn,trw->tn',RECC_System.ParameterDict['4_EI_WasteMgtEnergyIntensity'].Values[:,:,110,-1],RECC_System.FlowDict['F_9_10'].Values[:,:,:,0])
        SysVar_EnergyDemand_Remelting     = 1000 * np.einsum('mn,trm->tn',RECC_System.ParameterDict['4_EI_RemeltingEnergyIntensity'].Values[:,:,110,-1],RECC_System.FlowDict['F_9_12'].Values[:,:,:,0])
        # Unit: TJ/yr.
        
        # E) Calculate total energy demand
        SysVar_TotalEnergyDemand = np.einsum('trgn->tn',SysVar_EnergyDemand_UsePhase_Buildings_ByEnergyCarrier + SysVar_EnergyDemand_UsePhase_Vehicles_ByEnergyCarrier) + np.einsum('tmn->tn',SysVar_EnergyDemand_PrimaryProd) + SysVar_EnergyDemand_Manufacturing + SysVar_EnergyDemand_WasteMgt + SysVar_EnergyDemand_Remelting
        # Unit: TJ/yr.
        
        # F) Calculate direct emissions
        SysVar_DirectEmissions_UsePhase_Buildings = 0.001 * np.einsum('Xn,trgn->Xtrg',RECC_System.ParameterDict['6_PR_DirectEmissions'].Values,SysVar_EnergyDemand_UsePhase_Buildings_ByEnergyCarrier)
        SysVar_DirectEmissions_UsePhase_Vehicles  = 0.001 * np.einsum('Xn,trgn->Xtrg',RECC_System.ParameterDict['6_PR_DirectEmissions'].Values,SysVar_EnergyDemand_UsePhase_Vehicles_ByEnergyCarrier)
        SysVar_DirectEmissions_PrimaryProd        = 0.001 * np.einsum('Xn,tmn->Xtm'  ,RECC_System.ParameterDict['6_PR_DirectEmissions'].Values,SysVar_EnergyDemand_PrimaryProd)
        SysVar_DirectEmissions_Manufacturing      = 0.001 * np.einsum('Xn,tn->Xt'    ,RECC_System.ParameterDict['6_PR_DirectEmissions'].Values,SysVar_EnergyDemand_Manufacturing)
        SysVar_DirectEmissions_WasteMgt           = 0.001 * np.einsum('Xn,tn->Xt'    ,RECC_System.ParameterDict['6_PR_DirectEmissions'].Values,SysVar_EnergyDemand_WasteMgt)
        SysVar_DirectEmissions_Remelting          = 0.001 * np.einsum('Xn,tn->Xt'    ,RECC_System.ParameterDict['6_PR_DirectEmissions'].Values,SysVar_EnergyDemand_Remelting)
        # Unit: Mt/yr. 1 kg/MJ = 1kt/TJ
        
        # G) Calculate process emissions
        SysVar_ProcessEmissions_PrimaryProd       = np.einsum('mXt,tm->Xt'    ,RECC_System.ParameterDict['4_PE_ProcessExtensions'].Values[:,:,-1,:,mS],RECC_System.FlowDict['F_3_4'].Values[:,:,0])
        SysVar_ProcessEmissions_PrimaryProd_m     = np.einsum('mXt,tm->Xtm'   ,RECC_System.ParameterDict['4_PE_ProcessExtensions'].Values[:,:,-1,:,mS],RECC_System.FlowDict['F_3_4'].Values[:,:,0])
        # Unit: Mt/yr.
        
        # H) Calculate emissions from energy supply
        SysVar_IndirectGHGEms_EnergySupply_UsePhase_Buildings      = 0.001 * np.einsum('Xnrt,trgn->Xt',RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply'].Values[:,:,mS,mR,:,:],SysVar_EnergyDemand_UsePhase_Buildings_ByEnergyCarrier)
        SysVar_IndirectGHGEms_EnergySupply_UsePhase_Vehicles       = 0.001 * np.einsum('Xnrt,trgn->Xt',RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply'].Values[:,:,mS,mR,:,:],SysVar_EnergyDemand_UsePhase_Vehicles_ByEnergyCarrier)
        SysVar_IndirectGHGEms_EnergySupply_PrimaryProd             = 0.001 * np.einsum('Xnt,tmn->Xt',  RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply'].Values[:,:,mS,mR,mr,:],SysVar_EnergyDemand_PrimaryProd)
        SysVar_IndirectGHGEms_EnergySupply_Manufacturing           = 0.001 * np.einsum('Xnt,tn->Xt',  RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply'].Values[:,:,mS,mR,mr,:],SysVar_EnergyDemand_Manufacturing)
        SysVar_IndirectGHGEms_EnergySupply_WasteMgt                = 0.001 * np.einsum('Xnt,tn->Xt',  RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply'].Values[:,:,mS,mR,mr,:],SysVar_EnergyDemand_WasteMgt)
        SysVar_IndirectGHGEms_EnergySupply_Remelting               = 0.001 * np.einsum('Xnt,tn->Xt',  RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply'].Values[:,:,mS,mR,mr,:],SysVar_EnergyDemand_Remelting)
        SysVar_IndirectGHGEms_EnergySupply_UsePhase_Buildings_EL   = 0.001 * np.einsum('Xrt,trg->Xt', RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply'].Values[:,0,mS,mR,:,:],SysVar_EnergyDemand_UsePhase_Buildings_ByEnergyCarrier[:,:,:,0]) # electricity only
        SysVar_IndirectGHGEms_EnergySupply_UsePhase_Vehicles_EL    = 0.001 * np.einsum('Xrt,trg->Xt', RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply'].Values[:,0,mS,mR,:,:],SysVar_EnergyDemand_UsePhase_Vehicles_ByEnergyCarrier[:,:,:,0])  # electricity only
        
        # Unit: Mt/yr.
        
        # Diagnostics:
        Aa = np.einsum('tcgrnV->trg',SysVar_EnergyDemand_UsePhase_Total)           # Total use phase energy demand
        Aa = np.einsum('tme->tme',Element_Material_Composition[:,:,:,mS,mR])       # Element composition over years
        Aa = np.einsum('tme->tme',Element_Material_Composition_raw[:,:,:,mS,mR])   # Element composition over years, with zero entries
        Aa = np.einsum('tgm->tgm',Manufacturing_Output[:,:,:,mS,mR])               # Manufacturing_output_by_material
        
        
        # Calibration
        E_Calib_Buildings = np.einsum('trgn->tr',SysVar_EnergyDemand_UsePhase_Buildings_ByEnergyCarrier[:,:,:,0:7])
        E_Calib_Vehicles  = np.einsum('trgn->tr',SysVar_EnergyDemand_UsePhase_Vehicles_ByEnergyCarrier[:,:,:,0:7])
        
        
        # Ha) Calculate emissions benefits
        if ScriptConfig['ScrapExportRecyclingCredit']:
            SysVar_EnergyDemand_RecyclingCredit                = -1 * 1000 * np.einsum('mn,tm->tmn'  ,RECC_System.ParameterDict['4_EI_ProcessEnergyIntensity'].Values[:,:,110,0],RECC_System.FlowDict['F_12_0'].Values[:,-1,:,0])
            SysVar_DirectEmissions_RecyclingCredit             = -1 * 0.001 * np.einsum('Xn,tmn->Xt' ,RECC_System.ParameterDict['6_PR_DirectEmissions'].Values,SysVar_EnergyDemand_RecyclingCredit)
            SysVar_ProcessEmissions_RecyclingCredit            = -1 * np.einsum('mXt,tm->Xt'         ,RECC_System.ParameterDict['4_PE_ProcessExtensions'].Values[:,:,-1,:,mS],RECC_System.FlowDict['F_12_0'].Values[:,-1,:,0])
            SysVar_IndirectGHGEms_EnergySupply_RecyclingCredit = -1 * 0.001 * np.einsum('Xnt,tmn->Xt',RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply'].Values[:,:,mS,mR,mr,:],SysVar_EnergyDemand_RecyclingCredit)
        else:
            SysVar_DirectEmissions_RecyclingCredit = 0
            SysVar_ProcessEmissions_RecyclingCredit = 0
            SysVar_IndirectGHGEms_EnergySupply_RecyclingCredit = 0
        
        # I) Calculate emissions of system, by process group
        SysVar_GHGEms_UsePhase            = np.einsum('Xtrg->Xt',SysVar_DirectEmissions_UsePhase_Buildings + SysVar_DirectEmissions_UsePhase_Vehicles)
        SysVar_GHGEms_UsePhase_Scope2_El  = SysVar_IndirectGHGEms_EnergySupply_UsePhase_Buildings_EL + SysVar_IndirectGHGEms_EnergySupply_UsePhase_Vehicles_EL
        SysVar_GHGEms_UsePhase_OtherIndir = SysVar_IndirectGHGEms_EnergySupply_UsePhase_Buildings + SysVar_IndirectGHGEms_EnergySupply_UsePhase_Vehicles - SysVar_GHGEms_UsePhase_Scope2_El
        SysVar_GHGEms_PrimaryMetal        = np.einsum('Xtm->Xt',SysVar_DirectEmissions_PrimaryProd) + SysVar_ProcessEmissions_PrimaryProd + SysVar_IndirectGHGEms_EnergySupply_PrimaryProd
        SysVar_GHGEms_Manufacturing       = SysVar_DirectEmissions_Manufacturing + SysVar_IndirectGHGEms_EnergySupply_Manufacturing
        SysVar_GHGEms_WasteMgtRemelting   = SysVar_DirectEmissions_WasteMgt + SysVar_DirectEmissions_Remelting + SysVar_IndirectGHGEms_EnergySupply_WasteMgt + SysVar_IndirectGHGEms_EnergySupply_Remelting
        SysVar_GHGEms_MaterialCycle       = SysVar_GHGEms_Manufacturing + SysVar_GHGEms_WasteMgtRemelting
                                             
        SysVar_GHGEms_RecyclingCredit     = SysVar_DirectEmissions_RecyclingCredit + SysVar_ProcessEmissions_RecyclingCredit + SysVar_IndirectGHGEms_EnergySupply_RecyclingCredit
        
        # J) Calculate total emissions of system
        SysVar_GHGEms_Other     = SysVar_GHGEms_UsePhase_Scope2_El + SysVar_GHGEms_UsePhase_OtherIndir + SysVar_GHGEms_PrimaryMetal + SysVar_GHGEms_MaterialCycle
        SysVar_TotalGHGEms      = SysVar_GHGEms_UsePhase + SysVar_GHGEms_Other + SysVar_GHGEms_RecyclingCredit

        SysVar_GHGEms_Materials = SysVar_GHGEms_PrimaryMetal + SysVar_GHGEms_MaterialCycle - SysVar_DirectEmissions_Manufacturing - SysVar_IndirectGHGEms_EnergySupply_Manufacturing

        # Unit: Mt/yr.
        
        # J) Calculate indicators
        SysVar_TotalGHGCosts     = np.einsum('t,Xt->Xt',RECC_System.ParameterDict['3_PR_RECC_CO2Price_SSP_32R'].Values[mR,:,mr,mS],SysVar_TotalGHGEms)
        # Unit: million $ / yr.
        
        # K) Compile results
        GHG_System[:,mS,mR]          = SysVar_TotalGHGEms[0,:].copy()
        GHG_UsePhase[:,mS,mR]        = SysVar_GHGEms_UsePhase[0,:].copy()
        GHG_UsePhase_Scope2_El[:,mS,mR]  = SysVar_GHGEms_UsePhase_Scope2_El[0,:].copy()
        GHG_UsePhase_OtherIndir[:,mS,mR] = SysVar_GHGEms_UsePhase_OtherIndir[0,:].copy()
        GHG_MaterialCycle[:,mS,mR]   = SysVar_GHGEms_MaterialCycle[0,:].copy()
        GHG_RecyclingCredit[:,mS,mR] = SysVar_GHGEms_RecyclingCredit[0,:].copy()
        GHG_Other[:,mS,mR]           = SysVar_GHGEms_Other[0,:].copy() # all non use-phase processes
        GHG_Materials[:,mS,mR]       = SysVar_GHGEms_Materials[0,:].copy()
        GHG_Vehicles[:,:,mS,mR]      = np.einsum('trg->tr',SysVar_DirectEmissions_UsePhase_Vehicles[0,:,:,:]).copy()
        GHG_Buildings[:,:,mS,mR]     = np.einsum('trg->tr',SysVar_DirectEmissions_UsePhase_Buildings[0,:,:,:]).copy()
        GHG_PrimaryMetal[:,mS,mR]    = SysVar_ProcessEmissions_PrimaryProd[0,:].copy() + SysVar_IndirectGHGEms_EnergySupply_PrimaryProd[0,:].copy() + np.einsum('Xtm->t',SysVar_DirectEmissions_PrimaryProd).copy()
        Material_Inflow[:,:,:,mS,mR] = np.einsum('trgm->tgm',RECC_System.FlowDict['F_6_7'].Values[:,:,:,:,0]).copy()
        Scrap_Outflow[:,:,mS,mR]     = np.einsum('trw->tw',RECC_System.FlowDict['F_9_10'].Values[:,:,:,0]).copy()
        PrimaryProduction[:,:,mS,mR] = RECC_System.FlowDict['F_3_4'].Values[:,:,0].copy()
        SecondaryProduct[:,:,mS,mR]  = RECC_System.FlowDict['F_9_12'].Values[:,-1,:,0].copy()
        FabricationScrap[:,:,mS,mR]  = RECC_System.FlowDict['F_5_10'].Values[:,-1,:,0].copy()
        GHG_Vehicles_id[:,mS,mR]     = SysVar_IndirectGHGEms_EnergySupply_UsePhase_Buildings[0,:].copy()
        GHG_Building_id[:,mS,mR]     = SysVar_IndirectGHGEms_EnergySupply_UsePhase_Vehicles[0,:].copy()
        GHG_Manufact_all[:,mS,mR]    = SysVar_GHGEms_Manufacturing[0,:].copy()
        GHG_WasteMgt_all[:,mS,mR]    = SysVar_GHGEms_WasteMgtRemelting[0,:].copy()
        EnergyCons_UP_Vh[:,mS,mR]    = np.einsum('trgn->t',SysVar_EnergyDemand_UsePhase_Vehicles_ByEnergyCarrier).copy()
        EnergyCons_UP_Bd[:,mS,mR]    = np.einsum('trgn->t',SysVar_EnergyDemand_UsePhase_Buildings_ByEnergyCarrier).copy()
        EnergyCons_UP_Mn[:,mS,mR]    = SysVar_EnergyDemand_Manufacturing.sum(axis =1).copy()
        EnergyCons_UP_Wm[:,mS,mR]    = SysVar_EnergyDemand_WasteMgt.sum(axis =1).copy() +  SysVar_EnergyDemand_Remelting.sum(axis =1).copy()
        Vehicle_km[:,mS,mR]          = np.einsum('tcgr->t',SysVar_StockServiceProvision_UsePhase[:,:,:,:,-1])
 
#Emissions scopes:
# System, all processes:         GHG_System        
# Use phase only:                GHG_UsePhase    
# Use phase, electricity scope2: GHG_UsePhase_Scope2_El
# Use phase, indirect, rest:     GHG_UsePhase_OtherIndir
# Primary metal production:      GHG_PrimaryMetal    
# Manufacturing_Recycling:       GHG_MaterialCycle
# RecyclingCredit                GHG_RecyclingCredit
        
        
# DIAGNOSTICS
a,b,c = RECC_System.Check_If_All_Chem_Elements_Are_present('F_5_6',0)
a,b,c = RECC_System.Check_If_All_Chem_Elements_Are_present('F_12_5',0)
a,b,c = RECC_System.Check_If_All_Chem_Elements_Are_present('F_5_10',0)
a,b,c = RECC_System.Check_If_All_Chem_Elements_Are_present('F_4_5',0)
a,b,c = RECC_System.Check_If_All_Chem_Elements_Are_present('F_17_6',0)

a,b,c = RECC_System.Check_If_All_Chem_Elements_Are_present('F_12_5',0)
a,b,c = RECC_System.Check_If_All_Chem_Elements_Are_present('F_12_0',0)
a,b,c = RECC_System.Check_If_All_Chem_Elements_Are_present('F_9_12',0)
        
#####################################################
#   Section 5) Evaluate results, save, and close    #
#####################################################
Mylog.info('## 5 - Evaluate results, save, and close')
### 5.1.) CREATE PLOTS and include them in log file
Mylog.info('### 5.1 - Create plots and include into logfiles')
Mylog.info('Plot and export results')

myfont = xlwt.Font()
myfont.bold = True
mystyle = xlwt.XFStyle()
mystyle.font = myfont
Result_workbook = xlwt.Workbook(encoding = 'ascii') # Export file
Sheet = Result_workbook.add_sheet('Cover')
Sheet.write(2,1,label = 'ScriptConfig', style = mystyle)
m = 3
for x in sorted(ScriptConfig.keys()):
    Sheet.write(m,1,label = x)
    Sheet.write(m,2,label = ScriptConfig[x])
    m +=1
Sheet = Result_workbook.add_sheet('Model_Results')    
ColLabels = ['Indicator','Unit','Region','Figure','RE scen','SocEc scen','ClimPol scen']
for m in range(0,len(ColLabels)):
    Sheet.write(0,m,label = ColLabels[m], style = mystyle)
for n in range(m+1,m+1+Nt):
    Sheet.write(0,n,label = int(IndexTable.Classification[IndexTable.index.get_loc('Time')].Items[n-m-1]), style = mystyle)

newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_System,1,len(ColLabels),'GHG emissions, system-wide','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'Figs 1 and 2','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_PrimaryMetal,newrowoffset,len(ColLabels),'GHG emissions, primary metal production','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'Figs 4 and 5','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,0.13 * PrimaryProduction[:,11,:,:],newrowoffset,len(ColLabels),'Cement production','Mt / yr',ScriptConfig['RegionalScope'],'Figs 9-11','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,PrimaryProduction[:,0:3,:,:].sum(axis=1),newrowoffset,len(ColLabels),'Primary steel production','Mt / yr',ScriptConfig['RegionalScope'],'Figs 6-8','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,SecondaryProduct.sum(axis=1),newrowoffset,len(ColLabels),'Secondary materials, total','Mt / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
for m in range(0,Nm):
    newrowoffset = msf.ExcelExportAdd_tAB(Sheet,Material_Inflow[:,:,m,:,:].sum(axis=1),newrowoffset,len(ColLabels),'Final consumption of materials: ' + IndexTable.Classification[IndexTable.index.get_loc('Engineering materials')].Items[m],'Mt / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
for m in range(0,Nw):
    newrowoffset = msf.ExcelExportAdd_tAB(Sheet,FabricationScrap[:,m,:,:],newrowoffset,len(ColLabels),'Fabrication scrap: ' + IndexTable.Classification[IndexTable.index.get_loc('Waste_Scrap')].Items[m],'Mt / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,SecondaryProduct[:,0:4,:,:].sum(axis=1),newrowoffset,len(ColLabels),'Secondary steel','Mt / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,SecondaryProduct[:,4:6,:,:].sum(axis=1),newrowoffset,len(ColLabels),'Secondary Al','Mt / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,SecondaryProduct[:,6,:,:],newrowoffset,len(ColLabels),'Secondary copper','Mt / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_UsePhase,newrowoffset,len(ColLabels),'GHG emissions, use phase','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_Other,newrowoffset,len(ColLabels),'GHG emissions, industries and energy supply','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)

newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_Vehicles.sum(axis =1),newrowoffset,len(ColLabels),'GHG emissions, vehicles, use phase','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_Buildings.sum(axis =1),newrowoffset,len(ColLabels),'GHG emissions, buildings, use phase','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_Vehicles_id,newrowoffset,len(ColLabels),'GHG emissions, vehicles, energy supply','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_Building_id,newrowoffset,len(ColLabels),'GHG emissions, buildings, energy supply','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_Manufact_all,newrowoffset,len(ColLabels),'GHG emissions, manufacturing, all','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_WasteMgt_all,newrowoffset,len(ColLabels),'GHG emissions, waste mgt. and remelting, all','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,PrimaryProduction[:,4:6,:,:].sum(axis=1),newrowoffset,len(ColLabels),'Primary Al production','Mt / yr',ScriptConfig['RegionalScope'],'Figs 6-8','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,PrimaryProduction[:,6,:,:],newrowoffset,len(ColLabels),'Primary Cu production','Mt / yr',ScriptConfig['RegionalScope'],'Figs 6-8','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_Materials,newrowoffset,len(ColLabels),'GHG emissions, material industries and their energy supply','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
# energy flows
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,EnergyCons_UP_Vh,newrowoffset,len(ColLabels),'Energy cons., use phase, vehicles','TJ/yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,EnergyCons_UP_Bd,newrowoffset,len(ColLabels),'Energy cons., use phase, buildings','TJ/yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,EnergyCons_UP_Mn,newrowoffset,len(ColLabels),'Energy cons., use phase, manufacturing','TJ/yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,EnergyCons_UP_Wm,newrowoffset,len(ColLabels),'Energy cons., use phase, waste mgt. and remelting','TJ/yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
# stocks
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,StockCurves_Totl[:,0,:,:],newrowoffset,len(ColLabels),'In-use stock, pass. vehicles','million units',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,StockCurves_Totl[:,1,:,:],newrowoffset,len(ColLabels),'In-use stock, res. buildings','million m2',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
for mg in range(0,Ng):
    newrowoffset = msf.ExcelExportAdd_tAB(Sheet,StockCurves_Prod[:,mg,:,:],newrowoffset,len(ColLabels),'In-use stock, ' + IndexTable.Classification[IndexTable.index.get_loc('Good')].Items[mg],'Vehicles: million, Buildings: million m2',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
#per capita stocks
for mr in range(0,Nr-1):
    for mG in range(0,NG):
        newrowoffset = msf.ExcelExportAdd_tAB(Sheet,pCStocksCurves[:,mG,mr,:,:],newrowoffset,len(ColLabels),'per capita in-use stock, ' + IndexTable.Classification[IndexTable.index.get_loc('Product Groups')].Items[mG],'vehicles: cars per person, buildings: m2 per person',IndexTable.Classification[IndexTable.index.get_loc('Region')].Items[mr],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
#population
for mr in range(0,Nr-1):
    newrowoffset = msf.ExcelExportAdd_tAB(Sheet,Population[:,mr,:,:],newrowoffset,len(ColLabels),'Population','million',IndexTable.Classification[IndexTable.index.get_loc('Region')].Items[mr],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
#Downsizing and Mat subst. shares
for mr in range(0,Nr-1):
    newrowoffset = msf.ExcelExportAdd_tAB(Sheet,np.einsum('tS,R->tSR',ParameterDict['3_SHA_DownSizing_Vehicles'].Values[0,mr,:,:],np.ones((NR))),newrowoffset,len(ColLabels),'Share of downsized pass. vehicles','%',IndexTable.Classification[IndexTable.index.get_loc('Region')].Items[mr],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)    
    newrowoffset = msf.ExcelExportAdd_tAB(Sheet,np.einsum('tS,R->tSR',ParameterDict['3_SHA_DownSizing_Buildings'].Values[0,mr,:,:],np.ones((NR))),newrowoffset,len(ColLabels),'Share of downsized res. buildings','%',IndexTable.Classification[IndexTable.index.get_loc('Region')].Items[mr],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)    
    for mg in range(0,6):
        newrowoffset = msf.ExcelExportAdd_tAB(Sheet,np.einsum('tS,R->tSR',ParameterDict['3_SHA_LightWeighting_Vehicles'].Values[mg,mr,:,:],np.ones((NR))),newrowoffset,len(ColLabels), 'Share of light-weighted ' +IndexTable.Classification[IndexTable.index.get_loc('Good')].Items[mg],'%',IndexTable.Classification[IndexTable.index.get_loc('Region')].Items[mr],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)    
    for mg in range(6,Ng):
        newrowoffset = msf.ExcelExportAdd_tAB(Sheet,np.einsum('tS,R->tSR',ParameterDict['3_SHA_LightWeighting_Buildings'].Values[mg,mr,:,:],np.ones((NR))),newrowoffset,len(ColLabels),'Share of light-weighted ' +IndexTable.Classification[IndexTable.index.get_loc('Good')].Items[mg],'%',IndexTable.Classification[IndexTable.index.get_loc('Region')].Items[mr],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)    
#vehicle km 
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,Vehicle_km,newrowoffset,len(ColLabels),'km driven by pass. vehicles','million km/yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
# Use phase indirect GHG, primary prodution GHG, material cycle and recycling credit
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_UsePhase_Scope2_El,newrowoffset,len(ColLabels),'GHG emissions, use phase scope 2 (electricity)','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_UsePhase_OtherIndir,newrowoffset,len(ColLabels),'GHG emissions, use phase other indirect (non-el.)','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_PrimaryMetal,newrowoffset,len(ColLabels),'GHG emissions, primary metal production','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_MaterialCycle,newrowoffset,len(ColLabels),'GHG emissions, manufact, wast mgt., remelting and indirect','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)
newrowoffset = msf.ExcelExportAdd_tAB(Sheet,GHG_RecyclingCredit,newrowoffset,len(ColLabels),'GHG emissions, recycling credits','Mt of CO2-eq / yr',ScriptConfig['RegionalScope'],'none','Cf. Cover sheet',IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items,IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items)


MyColorCycle = pylab.cm.Paired(np.arange(0,1,0.2))
#linewidth = [1.2,2.4,1.2,1.2,1.2]
linewidth  = [1.2,2,1.2]
linewidth2 = [1.2,2,1.2]

Figurecounter = 1
LegendItems_SSP    = IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items
LegendItems_RCP    = IndexTable.Classification[IndexTable.index.get_loc('Scenario_RCP')].Items
LegendItems_SSP_RE = ['LED, no EST', 'LED, 2°C ES', 'SSP1, no EST', 'SSP1, 2°C ES', 'SSP2, no EST', 'SSP2, 2°C ES']
LegendItems_SSP_UP = ['Use Phase, SSP1, no EST', 'Rest of system GHG, SSP1, no EST','Use Phase, SSP1, 2°C ES', 'Rest of system GHG, SSP1, 2°C ES']
ColorOrder         = [1,0,3]

# 1) Emissions by scenario

#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#for m in range(0,NS):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),GHG_System[:,m,0], linewidth = linewidth[m])
#    ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP),shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('GHG emissions of system, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('No energy system (ES) transformation (EST)', fontsize = 12) 
#if ScriptConfig['UseGivenPlotBoundaries'] == True:
#    plt.axis([2015, 2050, 0, ScriptConfig['Plot1Max']])
#plt.show()
#fig_name = 'GHG_No_EST'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#for m in range(0,NS):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),GHG_System[:,m,1], linewidth = linewidth[m])
#    ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP),shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('GHG emissions of system, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('With energy system (ES) transformation (EST)', fontsize = 12) 
#if ScriptConfig['UseGivenPlotBoundaries'] == True:
#    plt.axis([2015, 2050, 0, ScriptConfig['Plot1Max']])
#plt.show()
#fig_name = 'GHG_With_EST'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1

# policy baseline vs. RCP 2.6
fig1, ax1 = plt.subplots()
ax1.set_color_cycle(MyColorCycle)
ProxyHandlesList = []
for m in range(0,NS):
    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),GHG_System[:,m,0], linewidth = linewidth[m], color = MyColorCycle[ColorOrder[m],:])
    #ProxyHandlesList.append(plt.line((0, 0), 1, 1, fc=MyColorCycle[m,:]))
    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),GHG_System[:,m,1], linewidth = linewidth2[m], linestyle = '--', color = MyColorCycle[ColorOrder[m],:])
    #ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))     
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP_RE),shadow = False, prop={'size':9}, loc = 'lower left')# 'upper right' ,bbox_to_anchor=(1.20, 1))
plt_lgd  = plt.legend(LegendItems_SSP_RE,shadow = False, prop={'size':9}, loc = 'upper left')# 'upper right' ,bbox_to_anchor=(1.20, 1))    
plt.ylabel('GHG emissions of system, Mt/yr.', fontsize = 12) 
plt.xlabel('year', fontsize = 12) 
plt.title('System-wide emissions, by SSP scenario, '+ ScriptConfig['RegionalScope'] + '.', fontsize = 12) 
if ScriptConfig['UseGivenPlotBoundaries'] == True:
    plt.axis([2016, 2050, 0, ScriptConfig['Plot1Max']])
plt.show()
fig_name = 'GHG_Ems_Overview'
# include figure in logfile:
fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
Figurecounter += 1

# Primary material production emissions
fig1, ax1 = plt.subplots()
ax1.set_color_cycle(MyColorCycle)
ProxyHandlesList = []
for m in range(0,NS):
    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),GHG_PrimaryMetal[:,m,0])
plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
plt.ylabel('GHG emissions of primary material production, Mt/yr.', fontsize = 12) 
plt.xlabel('year', fontsize = 12) 
plt.title('GHG primary materials, no EST', fontsize = 12) 
if ScriptConfig['UseGivenPlotBoundaries'] == True:
    plt.axis([2015, 2050, 0, ScriptConfig['Plot2Max']])
plt.show()
fig_name = 'GHG_PP_NoEST'
# include figure in logfile:
fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
Figurecounter += 1

fig1, ax1 = plt.subplots()
ax1.set_color_cycle(MyColorCycle)
ProxyHandlesList = []
for m in range(0,NS):
    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),GHG_PrimaryMetal[:,m,1])
plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
plt.ylabel('GHG emissions of primary material production, Mt/yr.', fontsize = 12) 
plt.xlabel('year', fontsize = 12) 
plt.title('GHG primary materials, with EST', fontsize = 12) 
if ScriptConfig['UseGivenPlotBoundaries'] == True:
    plt.axis([2015, 2050, 0, ScriptConfig['Plot2Max']])
plt.show()
fig_name = 'GHG_PP_WithEST'
# include figure in logfile:
fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
Figurecounter += 1

## Primary material production
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#for m in range(0,NS):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), PrimaryProduction[:,0:3,m,0].sum(axis=1))
#plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('Primary steel production, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('Primary steel, no EST', fontsize = 12) 
#if ScriptConfig['UseGivenPlotBoundaries'] == True:    
#    plt.axis([2015, 2050, 0, 0.15 * ScriptConfig['Plot2Max']])
#plt.show()
#fig_name = 'PrimarySteel_NoEST'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#for m in range(0,NS):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), PrimaryProduction[:,0:3,m,1].sum(axis=1))
#plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('Primary steel production, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('Primary steel, with EST', fontsize = 12) 
#if ScriptConfig['UseGivenPlotBoundaries'] == True:
#    plt.axis([2015, 2050, 0, 0.15 * ScriptConfig['Plot2Max']])
#plt.show()
#fig_name = 'PrimarySteel_WithEST'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1

# primary steel, no CP and 2°C combined:
fig1, ax1 = plt.subplots()
ax1.set_color_cycle(MyColorCycle)
ProxyHandlesList = []
for m in range(0,NS):
    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),PrimaryProduction[:,0:3,m,0].sum(axis=1), linewidth = linewidth[m], color = MyColorCycle[ColorOrder[m],:])
    #ProxyHandlesList.append(plt.line((0, 0), 1, 1, fc=MyColorCycle[m,:]))
    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),PrimaryProduction[:,0:3,m,1].sum(axis=1), linewidth = linewidth2[m], linestyle = '--', color = MyColorCycle[ColorOrder[m],:])
    #ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))     
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP_RE),shadow = False, prop={'size':9}, loc = 'lower left')# 'upper right' ,bbox_to_anchor=(1.20, 1))
plt_lgd  = plt.legend(LegendItems_SSP_RE,shadow = False, prop={'size':9}, loc = 'upper right')# 'upper right' ,bbox_to_anchor=(1.20, 1))    
plt.ylabel('Primary steel production, Mt/yr.', fontsize = 12) 
plt.xlabel('year', fontsize = 12) 
plt.title('Primary steel production, by SSP scenario, '+ ScriptConfig['RegionalScope'] + '.', fontsize = 12) 
if ScriptConfig['UseGivenPlotBoundaries'] == True:
    plt.axis([2017, 2050, 0, 0.15 * ScriptConfig['Plot2Max']])
plt.show()
fig_name = 'PSteel_Overview'
# include figure in logfile:
fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
Figurecounter += 1


#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#for m in range(0,NS):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), 0.13 * PrimaryProduction[:,11,m,0]) # 0.13 is the cement content of concrete
#plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('Cement production, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('Cement production, no EST', fontsize = 12) 
#if ScriptConfig['UseGivenPlotBoundaries'] == True:
#    plt.axis([2020, 2050, 0, 0.3 * ScriptConfig['Plot2Max']])
#plt.show()
#fig_name = 'Cement_NoEST'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#for m in range(0,NS):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), 0.13 * PrimaryProduction[:,11,m,1]) # 0.13 is the cement content of concrete
#plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('Cement production, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('Cement production, with EST', fontsize = 12) 
#if ScriptConfig['UseGivenPlotBoundaries'] == True:
#    plt.axis([2020, 2050, 0, 0.3 * ScriptConfig['Plot2Max']])
#plt.show()
#fig_name = 'Cement_WithEST'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1

# Cement production, no RE and RE combined:
fig1, ax1 = plt.subplots()
ax1.set_color_cycle(MyColorCycle)
ProxyHandlesList = []
for m in range(0,NS):
    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),0.13 * PrimaryProduction[:,11,m,0], linewidth = linewidth[m], color = MyColorCycle[ColorOrder[m],:])
    #ProxyHandlesList.append(plt.line((0, 0), 1, 1, fc=MyColorCycle[m,:]))
    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),0.13 * PrimaryProduction[:,11,m,1], linewidth = linewidth2[m], linestyle = '--', color = MyColorCycle[ColorOrder[m],:])
    #ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))     
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP_RE),shadow = False, prop={'size':9}, loc = 'lower left')# 'upper right' ,bbox_to_anchor=(1.20, 1))
plt_lgd  = plt.legend(LegendItems_SSP_RE,shadow = False, prop={'size':9}, loc = 'upper right')# 'upper right' ,bbox_to_anchor=(1.20, 1))    
plt.ylabel('Cement production, Mt/yr.', fontsize = 12) 
plt.xlabel('year', fontsize = 12) 
plt.title('Cement production, by SSP scenario, '+ ScriptConfig['RegionalScope'] + '.', fontsize = 12) 
if ScriptConfig['UseGivenPlotBoundaries'] == True:
    plt.axis([2017, 2050, 0, 0.30 * ScriptConfig['Plot2Max']])
plt.show()
fig_name = 'Cement_Overview'
# include figure in logfile:
fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
Figurecounter += 1

## Plot on recycled steel
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#for m in range(0,NS):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), SecondaryProduct[:,0:4,m,0].sum(axis =1)) # all steel types
#plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('Secondary steel, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('Recycled steel, no EST', fontsize = 12) 
#if ScriptConfig['UseGivenPlotBoundaries'] == True:
#    plt.axis([2020, 2050, 0, 0.6* ScriptConfig['Plot3Max']])
#plt.show()
#fig_name = 'SecondarySteel_NoEST'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#for m in range(0,NS):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), SecondaryProduct[:,0:4,m,1].sum(axis =1)) # all steel types
#plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('Secondary steel, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('SecondarySteel_WithEST', fontsize = 12) 
#if ScriptConfig['UseGivenPlotBoundaries'] == True:
#    plt.axis([2020, 2050, 0, 0.6 * ScriptConfig['Plot3Max']])
#plt.show()
#fig_name = 'SecondarySteel_WithEST'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1

# Recycled steel, RE and no RE
fig1, ax1 = plt.subplots()
ax1.set_color_cycle(MyColorCycle)
ProxyHandlesList = []
for m in range(0,NS):
    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), SecondaryProduct[:,0:4,m,0].sum(axis =1), linewidth = linewidth[m], color = MyColorCycle[ColorOrder[m],:])
    #ProxyHandlesList.append(plt.line((0, 0), 1, 1, fc=MyColorCycle[m,:]))
    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), SecondaryProduct[:,0:4,m,1].sum(axis =1), linewidth = linewidth2[m], linestyle = '--', color = MyColorCycle[ColorOrder[m],:])
    #ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))     
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP_RE),shadow = False, prop={'size':9}, loc = 'lower left')# 'upper right' ,bbox_to_anchor=(1.20, 1))
plt_lgd  = plt.legend(LegendItems_SSP_RE,shadow = False, prop={'size':9}, loc = 'upper left')# 'upper right' ,bbox_to_anchor=(1.20, 1))    
plt.ylabel('Recycled steel and iron, Mt/yr.', fontsize = 12) 
plt.xlabel('year', fontsize = 12) 
plt.title('Recycled iron and steel, by SSP scenario, '+ ScriptConfig['RegionalScope'] + '.', fontsize = 12) 
if ScriptConfig['UseGivenPlotBoundaries'] == True:
    plt.axis([2018, 2050, 0, 0.8 * ScriptConfig['Plot3Max']])
plt.show()
fig_name = 'SteelRecycling_Overview'
# include figure in logfile:
fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
Figurecounter += 1

# Recycled Al, RE and no RE
fig1, ax1 = plt.subplots()
ax1.set_color_cycle(MyColorCycle)
ProxyHandlesList = []
for m in range(0,NS):
    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), SecondaryProduct[:,4:6,m,0].sum(axis =1), linewidth = linewidth[m], color = MyColorCycle[ColorOrder[m],:])
    #ProxyHandlesList.append(plt.line((0, 0), 1, 1, fc=MyColorCycle[m,:]))
    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), SecondaryProduct[:,4:6,m,1].sum(axis =1), linewidth = linewidth2[m], linestyle = '--', color = MyColorCycle[ColorOrder[m],:])
    #ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))     
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP_RE),shadow = False, prop={'size':9}, loc = 'lower left')# 'upper right' ,bbox_to_anchor=(1.20, 1))
plt_lgd  = plt.legend(LegendItems_SSP_RE,shadow = False, prop={'size':9}, loc = 'upper left')# 'upper right' ,bbox_to_anchor=(1.20, 1))    
plt.ylabel('Recycled aluminium, Mt/yr.', fontsize = 12) 
plt.xlabel('year', fontsize = 12) 
plt.title('Recycled aluminium, by SSP scenario, '+ ScriptConfig['RegionalScope'] + '.', fontsize = 12) 
if ScriptConfig['UseGivenPlotBoundaries'] == True:
    plt.axis([2018, 2050, 0, 0.06 * ScriptConfig['Plot3Max']])
plt.show()
fig_name = 'AluminiumRecycling_Overview'
# include figure in logfile:
fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
Figurecounter += 1

# Recycled copper, RE and no RE
fig1, ax1 = plt.subplots()
ax1.set_color_cycle(MyColorCycle)
ProxyHandlesList = []
for m in range(0,NS):
    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), SecondaryProduct[:,6,m,0], linewidth = linewidth[m], color = MyColorCycle[ColorOrder[m],:])
    #ProxyHandlesList.append(plt.line((0, 0), 1, 1, fc=MyColorCycle[m,:]))
    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), SecondaryProduct[:,6,m,1], linewidth = linewidth2[m], linestyle = '--', color = MyColorCycle[ColorOrder[m],:])
    #ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))     
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP_RE),shadow = False, prop={'size':9}, loc = 'lower left')# 'upper right' ,bbox_to_anchor=(1.20, 1))
plt_lgd  = plt.legend(LegendItems_SSP_RE,shadow = False, prop={'size':9}, loc = 'upper left')# 'upper right' ,bbox_to_anchor=(1.20, 1))    
plt.ylabel('Recycled copper, Mt/yr.', fontsize = 12) 
plt.xlabel('year', fontsize = 12) 
plt.title('Recycled copper, by SSP scenario, '+ ScriptConfig['RegionalScope'] + '.', fontsize = 12) 
if ScriptConfig['UseGivenPlotBoundaries'] == True:
    plt.axis([2018, 2050, 0, 0.04 * ScriptConfig['Plot3Max']])
plt.show()
fig_name = 'CopperRecycling_Overview'
# include figure in logfile:
fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
Figurecounter += 1


# Use phase and indirect emissions, RE and no RE
fig1, ax1 = plt.subplots()
ax1.set_color_cycle(MyColorCycle)
ProxyHandlesList = []
# Use phase and other ems., SSP1, no RE
ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), GHG_UsePhase[:,0,0] , linewidth = linewidth[2], color = MyColorCycle[ColorOrder[2],:])
ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), GHG_Other[:,0,0] ,    linewidth = linewidth[2], color = MyColorCycle[ColorOrder[2],:], linestyle = '--')
# Use phase and other ems., SSP1, with RE
ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), GHG_UsePhase[:,0,1] , linewidth = linewidth[2], color = MyColorCycle[ColorOrder[1],:])
ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), GHG_Other[:,0,1] ,    linewidth = linewidth[2], color = MyColorCycle[ColorOrder[1],:], linestyle = '--')
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP_RE),shadow = False, prop={'size':9}, loc = 'lower left')# 'upper right' ,bbox_to_anchor=(1.20, 1))
plt_lgd  = plt.legend(LegendItems_SSP_UP,shadow = False, prop={'size':9}, loc = 'upper right')# 'upper right' ,bbox_to_anchor=(1.20, 1))    
plt.ylabel('GHG emissions, Mt/yr.', fontsize = 12) 
plt.xlabel('year', fontsize = 12) 
plt.title('GHG emissions by process and scenario, SSP1, '+ ScriptConfig['RegionalScope'] + '.', fontsize = 12) 
if ScriptConfig['UseGivenPlotBoundaries'] == True:
    plt.axis([2018, 2050, 0, 0.75 * ScriptConfig['Plot1Max']])
plt.show()
fig_name = 'GHG_UsePhase_Overview'
# include figure in logfile:
fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
Figurecounter += 1

# Plot implementation curves
fig1, ax1 = plt.subplots()
ax1.set_color_cycle(MyColorCycle)
ProxyHandlesList = []
for m in range(0,NR):
    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1), RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp'].Values[:,-1,m]) # first region
plt_lgd  = plt.legend(LegendItems_RCP,shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
plt.ylabel('Stock reduction potential seized, %.', fontsize = 12) 
plt.xlabel('year', fontsize = 12) 
plt.title('Implementation curves for more intense use, by region and scenario', fontsize = 12) 
if ScriptConfig['UseGivenPlotBoundaries'] == True:    
    plt.axis([2020, 2050, 0, 110])
plt.show()
fig_name = 'ImplementationCurves_' + IndexTable.Classification[IndexTable.index.get_loc('Region')].Items[0]
# include figure in logfile:
fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
Figurecounter += 1


# Plot system emissions, by process, stacked.
# Area plot, stacked, GHG emissions, material production, waste mgt, remelting, etc.
MyColorCycle = pylab.cm.gist_earth(np.arange(0,1,0.155)) # select 12 colors from the 'Set1' color map.            
grey0_9      = np.array([0.9,0.9,0.9,1])

SSPScens   = ['LED','SSP1','SSP2']
RCPScens   = ['No climate policy','2 degrees C energy mix']
Area       = ['use phase','use phase, scope 2 (el)','use phase, other indirect','primary metal product.','manufact. & recycling','total (incl. recycling credit)']     

for mS in range(0,NS): # SSP
    for mR in range(0,NR): # RCP

        fig  = plt.figure(figsize=(8,5))
        ax1  = plt.axes([0.08,0.08,0.85,0.9])
        
        ProxyHandlesList = []   # For legend     
        
        # plot area
        ax1.fill_between(np.arange(2015,2061),np.zeros((Nt)), GHG_UsePhase[:,mS,mR], linestyle = '-', facecolor = MyColorCycle[1,:], linewidth = 0.5)
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[1,:])) # create proxy artist for legend
        ax1.fill_between(np.arange(2015,2061),GHG_UsePhase[:,mS,mR], GHG_UsePhase[:,mS,mR] + GHG_UsePhase_Scope2_El[:,mS,mR], linestyle = '-', facecolor = MyColorCycle[2,:], linewidth = 0.5)
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[2,:])) # create proxy artist for legend
        ax1.fill_between(np.arange(2015,2061),GHG_UsePhase[:,mS,mR] + GHG_UsePhase_Scope2_El[:,mS,mR], GHG_UsePhase[:,mS,mR] + GHG_UsePhase_Scope2_El[:,mS,mR] + GHG_UsePhase_OtherIndir[:,mS,mR], linestyle = '-', facecolor = MyColorCycle[3,:], linewidth = 0.5)
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[3,:])) # create proxy artist for legend
        ax1.fill_between(np.arange(2016,2061),GHG_UsePhase[1::,mS,mR] + GHG_UsePhase_Scope2_El[1::,mS,mR] + GHG_UsePhase_OtherIndir[1::,mS,mR], GHG_UsePhase[1::,mS,mR] + GHG_UsePhase_Scope2_El[1::,mS,mR] + GHG_UsePhase_OtherIndir[1::,mS,mR] + GHG_PrimaryMetal[1::,mS,mR], linestyle = '-', facecolor = MyColorCycle[4,:], linewidth = 0.5)
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[4,:])) # create proxy artist for legend    
        ax1.fill_between(np.arange(2016,2061),GHG_UsePhase[1::,mS,mR] + GHG_UsePhase_Scope2_El[1::,mS,mR] + GHG_UsePhase_OtherIndir[1::,mS,mR] + GHG_PrimaryMetal[1::,mS,mR], GHG_UsePhase[1::,mS,mR] + GHG_UsePhase_Scope2_El[1::,mS,mR] + GHG_UsePhase_OtherIndir[1::,mS,mR] + GHG_PrimaryMetal[1::,mS,mR] + GHG_MaterialCycle[1::,mS,mR], linestyle = '-', facecolor = MyColorCycle[5,:], linewidth = 0.5)
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[5,:])) # create proxy artist for legend    
        plt.plot(np.arange(2016,2061), GHG_System[1::,mS,mR] , linewidth = linewidth[2], color = 'k')
        plta = Line2D(np.arange(2016,2061), GHG_System[1::,mS,mR] , linewidth = linewidth[2], color = 'k')
        ProxyHandlesList.append(plta) # create proxy artist for legend    
        #plt.text(Data[m,:].min()*0.55, 7.8, 'Baseline: ' + ("%3.0f" % Base[m]) + ' Mt/yr.',fontsize=14,fontweight='bold')
        
        plt.title('GHG emissions, stacked by process group, \n' + ScriptConfig['RegionalScope'] + ', ' + SSPScens[mS] + ', ' + RCPScens[mR] + '.', fontsize = 18)
        plt.ylabel('Mt of CO2-eq.', fontsize = 18)
        plt.xlabel('Year', fontsize = 18)
        plt.xticks(fontsize=18)
        plt.yticks(fontsize=18)
        plt_lgd  = plt.legend(handles = reversed(ProxyHandlesList),labels = reversed(Area), shadow = False, prop={'size':14},ncol=1, loc = 'upper right')# ,bbox_to_anchor=(1.91, 1)) 
        ax1.set_xlim([2015, 2050])
        
        plt.show()
        fig_name = 'GHG_TimeSeries_AllProcesses_Stacked_' + ScriptConfig['RegionalScope'] + ', ' + SSPScens[mS] + ', ' + RCPScens[mR] + '.png'
        # include figure in logfile:
        fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
        fig.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=300, bbox_inches='tight')
        Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
        Figurecounter += 1

# Area plot, for material industries:
Area2   = ['primary metal product.','waste mgt. & recycling','manufacturing']     

for mS in range(0,NS): # SSP
    for mR in range(0,NR): # RCP

        fig  = plt.figure(figsize=(8,5))
        ax1  = plt.axes([0.08,0.08,0.85,0.9])
        
        ProxyHandlesList = []   # For legend     
        
        # plot area
        ax1.fill_between(np.arange(2016,2061),np.zeros((Nt-1)), GHG_PrimaryMetal[1::,mS,mR], linestyle = '-', facecolor = MyColorCycle[4,:], linewidth = 0.5)
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[4,:])) # create proxy artist for legend
        ax1.fill_between(np.arange(2016,2061),GHG_PrimaryMetal[1::,mS,mR], GHG_PrimaryMetal[1::,mS,mR] + GHG_WasteMgt_all[1::,mS,mR], linestyle = '-', facecolor = MyColorCycle[5,:], linewidth = 0.5)
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[5,:])) # create proxy artist for legend
        ax1.fill_between(np.arange(2016,2061),GHG_PrimaryMetal[1::,mS,mR] + GHG_WasteMgt_all[1::,mS,mR], GHG_PrimaryMetal[1::,mS,mR] + GHG_WasteMgt_all[1::,mS,mR] + GHG_Manufact_all[1::,mS,mR], linestyle = '-', facecolor = MyColorCycle[6,:], linewidth = 0.5)
        ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[6,:])) # create proxy artist for legend
        
        
        plt.title('GHG emissions, stacked by process group, \n' + ScriptConfig['RegionalScope'] + ', ' + SSPScens[mS] + ', ' + RCPScens[mR] + '.', fontsize = 18)
        plt.ylabel('Mt of CO2-eq.', fontsize = 18)
        plt.xlabel('Year', fontsize = 18)
        plt.xticks(fontsize=18)
        plt.yticks(fontsize=18)
        plt_lgd  = plt.legend(handles = reversed(ProxyHandlesList),labels = reversed(Area2), shadow = False, prop={'size':14},ncol=1, loc = 'upper right')# ,bbox_to_anchor=(1.91, 1)) 
        ax1.set_xlim([2015, 2050])
        
        plt.show()
        fig_name = 'GHG_TimeSeries_Materials_Stacked_' + ScriptConfig['RegionalScope'] + ', ' + SSPScens[mS] + ', ' + RCPScens[mR] + '.png'
        # include figure in logfile:
        fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name + '_' + ScriptConfig['RegionalScope'] + '.png'
        fig.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=300, bbox_inches='tight')
        Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
        Figurecounter += 1


### 5.2) Export to Excel
Mylog.info('### 5.2 - Export to Excel')
# Export list data
Result_workbook.save(os.path.join(ProjectSpecs_Path_Result,'ODYM_RECC_ModelResults_'+ ScriptConfig['Current_UUID'] + '.xls'))

# Export table data
Result_workbook = xlwt.Workbook(encoding = 'ascii') # Export element stock by region

Sheet = Result_workbook.add_sheet('Cover')
Sheet.write(2,1,label = 'ScriptConfig', style = mystyle)
m = 3
for x in sorted(ScriptConfig.keys()):
    Sheet.write(m,1,label = x)
    Sheet.write(m,2,label = ScriptConfig[x])
    m +=1

MyLabels= []
for S in range(0,NS):
    for R in range(0,NR):
        MyLabels.append(RECC_System.IndexTable.set_index('IndexLetter').loc['S'].Classification.Items[S] + ', ' + RECC_System.IndexTable.set_index('IndexLetter').loc['R'].Classification.Items[R])
    
ResultArray = GHG_System.reshape(Nt,NS * NR)    
msf.ExcelSheetFill(Result_workbook, 'TotalGHGFootprint', ResultArray, topcornerlabel = 'System-wide GHG emissions, Mt/yr', rowlabels = RECC_System.IndexTable.set_index('IndexLetter').loc['t'].Classification.Items, collabels = MyLabels, Style = mystyle, rowselect = None, colselect = None)

Result_workbook.save(os.path.join(ProjectSpecs_Path_Result,'SysVar_TotalGHGFootprint.xls'))

## 5.3) Export as .mat file
#Mylog.info('### 5.4 - Export to Matlab')
#Mylog.info('Saving stock data to Matlab.')
#Filestring_Matlab_out = os.path.join(ProjectSpecs_Path_Result, 'StockData.mat')
#scipy.io.savemat(Filestring_Matlab_out, mdict={'F_6_7_tgmSR_Mt/yr': Material_Inflow, 'F_9_10_twSR_Mt/yr': Scrap_Outflow})



### 5.4) Model run is finished. Wrap up.
Mylog.info('### 5.5 - Finishing')
Mylog.debug("Converting " + os.path.join(ProjectSpecs_Path_Result, '..', log_filename))
# everything from here on will not be included in the converted log file
msf.convert_log(os.path.join(ProjectSpecs_Path_Result, log_filename))
Mylog.info('Script is finished. Terminating logging process and closing all log files.')
Time_End = time.time()
Time_Duration = Time_End - Time_Start
Mylog.info('End of simulation: ' + time.asctime())
Mylog.info('Duration of simulation: %.1f seconds.' % Time_Duration)

# remove all handlers from logger
root = log.getLogger()
root.handlers = []  # required if you don't want to exit the shell
log.shutdown()

print('done.')


# The End.
