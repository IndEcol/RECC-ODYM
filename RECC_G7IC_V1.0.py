# -*- coding: utf-8 -*-
"""
Created on July 22, 2018

@authors: spauliuk
"""

"""
File RECC_G7_V1

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
import scipy
import pandas as pd
import shutil   
import uuid
import matplotlib.pyplot as plt   
import importlib
import getpass
from copy import deepcopy
from tqdm import tqdm
from scipy.interpolate import interp1d
import pylab

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
#Copy Config file into that folder
shutil.copy(ProjectSpecs_Name_ConFile, os.path.join(ProjectSpecs_Path_Result, ProjectSpecs_Name_ConFile))


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
#IndexTable.ix['t']['Classification'].Items # get classification content
Mylog.info('Read model data and parameters.')

ParameterDict = {}
mo_start = 0 # set mo for re-reading a certain parameter
for mo in range(mo_start,len(PL_Names)):
    #mo = 30 # set mo for re-reading a certain parameter
    #ParPath = os.path.join(os.path.abspath(os.path.join(ProjectSpecs_Path_Main, '.')), 'ODYM_RECC_Database', PL_Version[mo])
    ParPath = os.path.join(RECC_Paths.data_path, PL_Version[mo])
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

# Interpolate missing parameter values:

# 1) Material composition of vehicles:
# Values are given every 5 years, we need all values in between.
index = PL_Names.index('3_MC_RECC_Vehicles_G7IC')
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

# 2) Material composition of buildings:
# Values are given every 5 years, we need all values in between.
index = PL_Names.index('3_MC_RECC_Buildings_G7IC')
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

# 3) Energy intensity of historic products:
#index = PL_Names.index('3_EI_Products_UsePhase_USA')
#ParameterDict[PL_Names[index]].Values[0:115,:,:,:,:] = np.tile(ParameterDict[PL_Names[index]].Values[115,:,:,:,:],(115,1,1,1,1))

# 4) GHG intensity of energy supply:
# Extrapolate 2050-2060 as 2050 values
index = PL_Names.index('4_PE_GHGIntensityEnergySupply')
ParameterDict[PL_Names[index]].Values[:,:,:,36::] = np.einsum('XnS,t->XnSt',ParameterDict[PL_Names[index]].Values[:,:,:,35],np.ones(10))
# Replicate global average to regions
GHGEnergySupplyFull = np.zeros((NX,Nn,NS,Nr,Nt)) # Define new parameter array
GHGEnergySupplyFull = np.einsum('XnSt,r->XnSrt',ParameterDict[PL_Names[index]].Values.copy(),np.ones(Nr))
# Add MESSAGEix results from CDLINKS project
indexM = PL_Names.index('4_PE_GHGIntensityEnergySupply_CDLINKS_MESSAGEix')
GHGEnergySupplyFull[:,0,:,:,:] = ParameterDict[PL_Names[indexM]].Values[:,0,:,:,:].copy()
ParameterDict[PL_Names[indexM]].Values = GHGEnergySupplyFull.copy()

# 5) Fabrication yield:
# Extrapolate 2050-2060 as 2050 values
index = PL_Names.index('4_PY_Manufacturing_USA')
ParameterDict[PL_Names[index]].Values[:,:,:,:,1::,:] = np.einsum('t,mwgFr->mwgFtr',np.ones(45),ParameterDict[PL_Names[index]].Values[:,:,:,:,0,:])

# 6 Model flow control: Include or exclude certain sectors
if ScriptConfig['SectorSelect'] == 'passenger vehicles':
    ParameterDict['2_S_RECC_FinalProducts_2015_G7IC'].Values[:,:,6::,:] = 0
    ParameterDict['2_S_RECC_FinalProducts_Future_G7IC'].Values[:,:,1,:] = 0
if ScriptConfig['SectorSelect'] == 'residential buildings':
    ParameterDict['2_S_RECC_FinalProducts_2015_G7IC'].Values[:,:,0:6,:] = 0
    ParameterDict['2_S_RECC_FinalProducts_Future_G7IC'].Values[:,:,0,:] = 0
 
##########################################################
#    Section 3) Initialize dynamic MFA model for RECC    #
##########################################################
Mylog.info('## 3 - Initialize dynamic MFA model for RECC')
Mylog.info('Define RECC system and processes.')

#  Examples for testing
mS = 1
mR = 1
mr = 0 # region for GHG prices and intensities

# Select and loop over scenarios
for mS in range(0,NS):
    for mR in range(0,NR):
        
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
        
        RECC_System.FlowDict['F_9_0'] = msc.Flow(Name='waste mgt. and remelting losses' , P_Start = 9, P_End = 0, 
                                                 Indices = 't,r,e', Values=None, Uncert=None, 
                                                 Color = None, ID = None, UUID = None)
        
        # Define system variables: Stocks.
        RECC_System.StockDict['S_7']   = msc.Stock(Name='In-use stock', P_Res=7, Type=1,
                                                 Indices = 't,c,r,g,m,e', Values=None, Uncert=None,
                                                 ID=None, UUID=None)
        
        RECC_System.StockDict['S_10']   = msc.Stock(Name='Fabrication scrap buffer', P_Res=10, Type=1,
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
        TotalStock_UsePhase_Hist_cgr = RECC_System.ParameterDict['2_S_RECC_FinalProducts_2015_G7IC'].Values[0,:,:,:]
        
        # Determine total future stock, product level. Units: Vehicles: million, Buildings: million m2.
        TotalStockCurves_UsePhase = np.einsum('tGr,tr->trG',RECC_System.ParameterDict['2_S_RECC_FinalProducts_Future_G7IC'].Values[mS,:,:,:],RECC_System.ParameterDict['2_P_RECC_Population_SSP_32R'].Values[0,:,:,mS]) # Here the population model M is set to its default and does not appear in the summation.
        
        # 2) Include (or not) the RE strategies for the use phase:
        
        # Include_REStrategy_MoreIntenseUse:
        if ScriptConfig['Include_REStrategy_MoreIntenseUse'] == 'True':
            TotalStockCurves_UsePhase = np.einsum('trg,trg->trg', (1 - np.einsum('tr,gr->trg',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp_SSP_32R'].Values[mR,:,:,mS]*0.01,RECC_System.ParameterDict['6_PR_MoreIntenseUse_USA'].Values[:,:,mS])),TotalStockCurves_UsePhase)    
        
        # Include_REStrategy_LifeTimeExtension: Product lifetime extension.
        # First, replicate lifetimes for all age-cohorts
        Par_RECC_ProductLifetime = np.einsum('c,gUr->grc',np.ones((Nc)),RECC_System.ParameterDict['3_LT_RECC_ProductLifetime'].Values) # Sums up over U, only possible because of 1:1 correspondence of U and g!
        # Second, change lifetime of future age-cohorts according to lifetime extension parameter
        # This is equation 10 of the paper:
        if ScriptConfig['Include_REStrategy_LifeTimeExtension'] == 'True':
            Par_RECC_ProductLifetime[:,:,SwitchTime -1::] = np.einsum('crg,grc->grc',1 + np.einsum('cr,gr->crg',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp_SSP_32R'].Values[mR,:,:,mS]*0.01,RECC_System.ParameterDict['6_PR_LifeTimeExtension_USA'].Values[:,:,mS]),Par_RECC_ProductLifetime[:,:,SwitchTime -1::])
        
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
                FutureStock     = TotalStockCurves_UsePhase[1::, r, G]# Future total stock
                InitialStock    = TotalStock_UsePhase_Hist_cgr[:,:,r].copy()
                if G == 0: # quick fix for vehicles and buildings only !!!
                    InitialStock[:,6::] = 0  # set not relevant initial stock to 0
                if G == 1:
                    InitialStock[:,0:6] = 0  # set not relevant initial stock to 0
                SFArrayCombined = SF_Array[:,:,:,r]
                TypeSplit       = RECC_System.ParameterDict['3_SHA_TypeSplit_NewProducts_USA'].Values[G,:,1::,r,mS].transpose()
  
                Var_S, Var_O, Var_I = msf.compute_stock_driven_model_initialstock_typesplit(FutureStock,InitialStock,SFArrayCombined,TypeSplit, NegativeInflowCorrect = False)
                
                Stock_Detail_UsePhase[1::,:,:,r]   += Var_S.copy() # tcgr
                Outflow_Detail_UsePhase[1::,:,:,r] += Var_O.copy() # tcgr
                Inflow_Detail_UsePhase[1::,:,r]    += Var_I[SwitchTime::,:].copy() # tgr

        # Here so far: Units: Vehicles: million, Buildings: million m2. for stocks, X/yr for flows.
        
        # Clean up
        #del TotalStockCurves_UsePhase
        #del SF_Array
        
        # Prepare parameters:        
        # include light-weighting in future MC parameter
        Par_RECC_MC = RECC_System.ParameterDict['3_MC_RECC_Buildings_G7IC'].Values[:,1,:,:,:] + RECC_System.ParameterDict['3_MC_RECC_Vehicles_G7IC'].Values[:,0,:,:,:]
        if ScriptConfig['Include_REStrategy_LightWeighting'] == 'True':
            Par_RECC_MC[SwitchTime-1::,:,:,:] = np.einsum('cmgr,crg->cmgr',Par_RECC_MC[SwitchTime-1::,:,:,:], 1 - np.einsum('cr,gr->crg',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp_SSP_32R'].Values[mR,:,:,mS]*0.01,RECC_System.ParameterDict['6_PR_LightWeighting_USA'].Values[:,:,mS])) # crgm
        # Units: Vehicles: kg/unit, Buildings: kg/m2  
        
        # historic element content of materials in producuts:
        Par_Element_Material_Composition_of_Products = np.zeros((Nc,Nr,Ng,Nm,Ne)) # crgme
        Par_Element_Material_Composition_of_Products[0:Nc-Nt+1,:,:,:,:] = np.einsum('cmgr,me->crgme',Par_RECC_MC[0:Nc-Nt+1,:,:,:],RECC_System.ParameterDict['3_MC_Elements_Materials_ExistingStock'].Values)
        # For future age-cohorts, this parameter will be updated year by year in the loop below.
    
        # Manufacturing yield:
        Par_FabYield = np.einsum('mwggtr->mwgtr',RECC_System.ParameterDict['4_PY_Manufacturing_USA'].Values) # take diagonal of product = manufacturing process
        # Consider Fabrication yield improvement
        if ScriptConfig['Include_REStrategy_FabYieldImprovement'] == 'True':
            Par_FabYieldImprovement = np.einsum('w,tmgr->mwgtr',np.ones((Nw)),np.einsum('tr,mgr->tmgr',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp_SSP_32R'].Values[mR,:,:,mS]*0.01,RECC_System.ParameterDict['6_PR_FabricationYieldImprovement_USA'].Values[:,:,:,mS]))
        else:
            Par_FabYieldImprovement = 0
        Par_FabYield_Raster = Par_FabYield > 0    
        Par_FabYield        = Par_FabYield - Par_FabYield_Raster * Par_FabYieldImprovement #mwgtr
        Par_FabYield_total  = np.einsum('mwgtr->mgtr',Par_FabYield)
        Par_FabYield_total_inv = 1/(1-Par_FabYield_total) # mgtr
        Par_FabYield_total_inv[Par_FabYield_total_inv == 1] = 0
            
        # Consider EoL recovery rate improvement:
        if ScriptConfig['Include_REStrategy_EoL_RR_Improvement'] == 'True':
            Par_RECC_EoL_RR =  np.einsum('t,grmw->trmgw',np.ones((Nt)),RECC_System.ParameterDict['4_PY_EoL_RR_USA'].Values[:,:,:,:,0] *0.01) \
            + np.einsum('tr,grmw->trmgw',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp_SSP_32R'].Values[mR,:,:,mS]*0.01,RECC_System.ParameterDict['6_PR_EoL_RR_Improvement_USA'].Values[:,:,:,:,0]*0.01)
        else:    
            Par_RECC_EoL_RR =  np.einsum('t,grmw->trmgw',np.ones((Nt)),RECC_System.ParameterDict['4_PY_EoL_RR_USA'].Values[:,:,:,:,0] *0.01)
        
        Mylog.info('Translate total flows into individual materials and elements, for 2015 and historic age-cohorts.')
        
        # 1) Inflow
        RECC_System.FlowDict['F_6_7'].Values[0,:,:,:,:]   = \
        np.einsum('rgme,gr->rgme',Par_Element_Material_Composition_of_Products[SwitchTime-1,:,:,:,:],Inflow_Detail_UsePhase[0,:,:])/1000 # all elements, Indices='t,r,g,m,e'

        #2) Outflow            
        RECC_System.FlowDict['F_7_8'].Values[0,0:SwitchTime,:,:,:,:] = \
        np.einsum('crgme,cgr->crgme',Par_Element_Material_Composition_of_Products[0:SwitchTime,:,:,:,:],Outflow_Detail_UsePhase[0,0:SwitchTime,:,:])/1000 # all elements, Indices='t,r,g,m,e'
        
        #3) Stock
        RECC_System.StockDict['S_7'].Values[0,0:SwitchTime,:,:,:,:] = \
        np.einsum('trgme,cgr->crgme',Par_Element_Material_Composition_of_Products[0:SwitchTime,:,:,:,:],Stock_Detail_UsePhase[0,0:SwitchTime,:,:])/1000 # all elements, Indices='t,r,g,m,e'
        #Units so far: Mt/yr
        
        Mylog.info(' Calculate material stocks and flows, material cycles, determine elemental composition.')
        # Units: Mt and Mt/yr.
        # This calculation is done year-by-year, and the elemental composition of the materials is in part determined by the scrap flow metal composition
        for t in tqdm(range(1, Nt), unit=' years'): # 1: 2016
            CohortOffset = t +Nc -Nt # index of current age-cohort.   
            
            # 1) Split outflow into materials and chemical elements.
            RECC_System.FlowDict['F_7_8'].Values[t,0:CohortOffset,:,:,:,:] = \
            np.einsum('crgme,cgr->crgme',Par_Element_Material_Composition_of_Products[0:CohortOffset,:,:,:,:],Outflow_Detail_UsePhase[t,0:CohortOffset,:,:])/1000 # All elements.
            
            # 2) Calculate obsolete stock formation
            # None. 
            # RECC_System.FlowDict['F_8_0'].Values = 0. Already defined
            
            # 2) Consider re-use
            if ScriptConfig['Include_REStrategy_ReUse'] == 'True':
                RECC_System.FlowDict['F_8_17'].Values[t,0:CohortOffset,:,:,:,:] = np.einsum('rg,cme->crgme',np.einsum('r,gr->rg',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp_SSP_32R'].Values[mR,t,:,mS]*0.01,RECC_System.ParameterDict['6_PR_ReUse_USA'].Values[:,:,mS]),np.ones((CohortOffset,Nm,Ne))) * (RECC_System.FlowDict['F_7_8'].Values[t,0:CohortOffset,:,:,:,:] - RECC_System.FlowDict['F_8_0'].Values[t,0:CohortOffset,:,:,:,:])
                RECC_System.FlowDict['F_17_6'].Values[t,0:CohortOffset,:,:,:,:] = RECC_System.FlowDict['F_8_17'].Values[t,0:CohortOffset,:,:,:,:]
                # now, re-use only happens within the same region. Export to other regions needs to be added later.
                
            # 3) calculate inflow waste mgt.
            RECC_System.FlowDict['F_8_9'].Values[t,:,:,:,:]     = np.einsum('crgme->rgme',RECC_System.FlowDict['F_7_8'].Values[t,0:CohortOffset,:,:,:,:] - RECC_System.FlowDict['F_8_0'].Values[t,0:CohortOffset,:,:,:,:] - RECC_System.FlowDict['F_8_17'].Values[t,0:CohortOffset,:,:,:,:])
    
            # 4) EoL products to postconsumer scrap: trwe
            RECC_System.FlowDict['F_9_10'].Values[t,:,:,:]      = np.einsum('rmgw,rgme->rwe',Par_RECC_EoL_RR[t,:,:,:,:],RECC_System.FlowDict['F_8_9'].Values[t,:,:,:,:])    
            # Aggregate scrap flows at world level:
            RECC_System.FlowDict['F_9_10'].Values[t,-1,:,:]     = np.einsum('rme->me',RECC_System.FlowDict['F_9_10'].Values[t,0:-1,:,:].copy())
            RECC_System.FlowDict['F_9_10'].Values[t,0:-1,:,:]   = 0
        
            # 5) Add new scrap and calculate remelting.
            # Add old scrap with manufacturing scrap from last year. In year 2016, no fabrication scrap exists yet.
            RECC_System.FlowDict['F_10_9'].Values[t,-1,:,:]     = np.einsum('rwe->we',RECC_System.FlowDict['F_9_10'].Values[t,:,:,:] + RECC_System.StockDict['S_10'].Values[t-1,:,:,:].copy())
            RECC_System.StockDict['S_10'].Values[t-1,:,:,:]     = 0 # Fabriation scrap buffer is cleared
            RECC_System.FlowDict['F_9_12'].Values[t,:,:,:]      = np.einsum('rwe,wmePr->rme',RECC_System.FlowDict['F_10_9'].Values[t,:,:,:],RECC_System.ParameterDict['4_PY_MaterialProductionRemelting'].Values[:,:,:,:,0,:])
            RECC_System.FlowDict['F_12_5'].Values[t,:,:,:]      = RECC_System.FlowDict['F_9_12'].Values[t,:,:,:]
                        
            # 6) Waste mgt. losses.
            RECC_System.FlowDict['F_9_0'].Values[t,:,:]         = np.einsum('rgme->re',RECC_System.FlowDict['F_8_9'].Values[t,:,:,:,:]) + np.einsum('rwe->re',RECC_System.FlowDict['F_10_9'].Values[t,:,:,:]) - np.einsum('rwe->re',RECC_System.FlowDict['F_9_10'].Values[t,:,:,:]) - np.einsum('rme->re',RECC_System.FlowDict['F_9_12'].Values[t,:,:,:])
            
            # 7) Calculate manufacturing output, in Mt/yr, all elements, trgme, element composition not yet known.
            Manufacturing_Output_gm   = np.einsum('rgm->gm',RECC_System.FlowDict['F_6_7'].Values[t,:,:,:,0]) - np.einsum('crgm->gm',RECC_System.FlowDict['F_17_6'].Values[t,:,:,:,:,0])
        
            # 8) Calculate manufacturing input and primary production, all elements, element composition not yet known.
            Manufacturing_Input_m        = np.einsum('mg,gm->m',Par_FabYield_total_inv[:,:,t,-1],Manufacturing_Output_gm)
            Manufacturing_Input_gm       = np.einsum('mg,gm->gm',Par_FabYield_total_inv[:,:,t,-1],Manufacturing_Output_gm)
            Manufacturing_Input_Split_gm = np.einsum('gm,m->gm',Manufacturing_Input_gm, 1/Manufacturing_Input_m)
            Manufacturing_Input_Split_gm[np.isnan(Manufacturing_Input_Split_gm)] = 0
            
            PrimaryProduction_m          = Manufacturing_Input_m - RECC_System.FlowDict['F_12_5'].Values[t,-1,:,0]# secondary material comes first, no rebound! 
        
            RECC_System.FlowDict['F_4_5'].Values[t,:,:] = np.einsum('m,me->me',PrimaryProduction_m,RECC_System.ParameterDict['3_MC_Elements_Materials_Primary'].Values)
            RECC_System.FlowDict['F_3_4'].Values[t,:,:] = RECC_System.FlowDict['F_4_5'].Values[t,:,:]
            RECC_System.FlowDict['F_0_3'].Values[t,:,:] = RECC_System.FlowDict['F_3_4'].Values[t,:,:]
        
            Manufacturing_Input_me      = RECC_System.FlowDict['F_4_5'].Values[t,:,:] + RECC_System.FlowDict['F_12_5'].Values[t,-1,:,:] 
            Manufacturing_Input_gme     = np.einsum('me,gm->gme',Manufacturing_Input_me,Manufacturing_Input_Split_gm)       
        
            # 9) Calculate manufacturing scrap 
            RECC_System.FlowDict['F_5_10'].Values[t,:,:,:] = np.einsum('gme,mwgr->rwe',Manufacturing_Input_gme,Par_FabYield[:,:,:,t,:])
            # Fabrication scrap, to be recycled next year:
            RECC_System.StockDict['S_10'].Values[t,:,:,:]  = RECC_System.FlowDict['F_5_10'].Values[t,:,:,:]
        
            # 10) Calculate element composition of materials of current year
            Element_Material_Composition_t = np.einsum('me,m->me',RECC_System.FlowDict['F_4_5'].Values[t,:,:] + RECC_System.FlowDict['F_12_5'].Values[t,-1,:,:],1/np.einsum('me->m',RECC_System.FlowDict['F_4_5'].Values[t,:,:] + RECC_System.FlowDict['F_12_5'].Values[t,-1,:,:]))
            Element_Material_Composition_t[np.isnan(Element_Material_Composition_t)] = 0
            Par_Element_Material_Composition_of_Products[CohortOffset,:,:,:,:] = np.einsum('mgr,me->rgme',Par_RECC_MC[CohortOffset,:,:,:],Element_Material_Composition_t) # crgme
            
            # 11) Calculate manufacturing output
            RECC_System.FlowDict['F_5_6'].Values[t,-1,:,:,:] = np.einsum('gme,gm->gme',Par_Element_Material_Composition_of_Products[CohortOffset,-1,:,:,:],Manufacturing_Output_gm)
            
            # 12) Calculate element composition of final consumption and latest age-cohort in in-use stock
            RECC_System.FlowDict['F_6_7'].Values[t,:,:,:,:]   = \
            np.einsum('rgme,gr->rgme',Par_Element_Material_Composition_of_Products[CohortOffset,:,:,:,:],Inflow_Detail_UsePhase[t,:,:])/1000 # all elements, Indices='t,r,g,m,e'
            
            RECC_System.StockDict['S_7'].Values[t,0:CohortOffset +1,:,:,:,:] = \
            np.einsum('crgme,cgr->crgme',Par_Element_Material_Composition_of_Products[0:CohortOffset +1,:,:,:,:],Stock_Detail_UsePhase[t,0:CohortOffset +1,:,:])/1000 # All elements.
 
        
        # Check whether flow value arrays match their indices, etc.
        RECC_System.Consistency_Check() 
    
        # Determine Mass Balance
        Bal = RECC_System.MassBalance()
        
        # A) Calculate intensity of operation
        SysVar_StockServiceProvision_UsePhase = np.einsum('tgVr,tcgr->tcgrV',RECC_System.ParameterDict['3_IO_Vehicles_UsePhase_G7IC'].Values[:,:,:,:,mS], Stock_Detail_UsePhase) + np.einsum('cgVr,tcgr->tcgrV',RECC_System.ParameterDict['3_IO_Buildings_UsePhase_G7IC'].Values[:,:,:,:,mS], Stock_Detail_UsePhase)
        # Unit: million km/yr for vehicles, million m2 for buildings by three use types: heating, cooling, and DHW.
        
        # B) Calculate total operational energy use
        SysVar_EnergyDemand_UsePhase_Total  = np.einsum('cgVnr,tcgrV->tcgrnV',RECC_System.ParameterDict['3_EI_Products_UsePhase_G7IC'].Values[:,:,:,:,:,mS], SysVar_StockServiceProvision_UsePhase)
        # Unit: TJ/yr for both vehicles and buildings.
        
        # C) Translate 'all' energy carriers to specific ones, use phase
        SysVar_EnergyDemand_UsePhase_Buildings_ByEnergyCarrier = np.einsum('cgrVn,tcrgV->trgn',RECC_System.ParameterDict['3_SHA_EnergyCarrierSplit_Buildings_G7IC'].Values[:,:,:,:,:,mS],SysVar_EnergyDemand_UsePhase_Total[:,:,:,:,-1,:].copy())
        SysVar_EnergyDemand_UsePhase_Vehicles_ByEnergyCarrier  = np.einsum('cgrVn,tcrgV->trgn',RECC_System.ParameterDict['3_SHA_EnergyCarrierSplit_Vehicles_G7IC'].Values[:,:,:,:,:,mS] ,SysVar_EnergyDemand_UsePhase_Total[:,:,:,:,-1,:].copy())
        
        # D) Calculate energy demand of the other industries
        SysVar_EnergyDemand_PrimaryProd   = 1000 * np.einsum('mn,tm->tmn',RECC_System.ParameterDict['4_EI_ProcessEnergyIntensity_USA'].Values[:,:,110,0],RECC_System.FlowDict['F_3_4'].Values[:,:,0])
        SysVar_EnergyDemand_Manufacturing = 1000 * np.einsum('gnr,trgm->tn',RECC_System.ParameterDict['4_EI_ManufacturingEnergyIntensity_USA'].Values[:,:,110,:],RECC_System.FlowDict['F_5_6'].Values[:,:,:,:,0])
        SysVar_EnergyDemand_WasteMgt      = 1000 * np.einsum('gnr,trg->tn',RECC_System.ParameterDict['4_EI_WasteMgtEnergyIntensity'].Values[:,0,:,110,:],np.einsum('trgm->trg',RECC_System.FlowDict['F_8_9'].Values[:,:,:,:,0]))
        SysVar_EnergyDemand_Remelting     = 1000 * np.einsum('wnr,trw->tn',RECC_System.ParameterDict['4_EI_RemeltingEnergyIntensity'].Values[0:-1,:,110,:],RECC_System.FlowDict['F_10_9'].Values[:,:,:,0])
        # Unit: TJ/yr.
        
        # E) Calculate total energy demand
        SysVar_TotalEnergyDemand = np.einsum('trgn->tn',SysVar_EnergyDemand_UsePhase_Buildings_ByEnergyCarrier + SysVar_EnergyDemand_UsePhase_Vehicles_ByEnergyCarrier) + np.einsum('tmn->tn',SysVar_EnergyDemand_PrimaryProd) + SysVar_EnergyDemand_Manufacturing + SysVar_EnergyDemand_WasteMgt + SysVar_EnergyDemand_Remelting
        # Unit: TJ/yr.
        
        # F) Calculate direct emissions
        SysVar_DirectEmissions_UsePhase_Buildings = np.einsum('Xn,trgn->Xtrg',RECC_System.ParameterDict['6_PR_DirectEmissions'].Values,SysVar_EnergyDemand_UsePhase_Buildings_ByEnergyCarrier)
        SysVar_DirectEmissions_UsePhase_Vehicles  = np.einsum('Xn,trgn->Xtrg',RECC_System.ParameterDict['6_PR_DirectEmissions'].Values,SysVar_EnergyDemand_UsePhase_Vehicles_ByEnergyCarrier)
        SysVar_DirectEmissions_PrimaryProd        = np.einsum('Xn,tmn->Xtm'  ,RECC_System.ParameterDict['6_PR_DirectEmissions'].Values,SysVar_EnergyDemand_PrimaryProd)
        SysVar_DirectEmissions_Manufacturing      = np.einsum('Xn,tn->Xt'    ,RECC_System.ParameterDict['6_PR_DirectEmissions'].Values,SysVar_EnergyDemand_Manufacturing)
        SysVar_DirectEmissions_WasteMgt           = np.einsum('Xn,tn->Xt'    ,RECC_System.ParameterDict['6_PR_DirectEmissions'].Values,SysVar_EnergyDemand_WasteMgt)
        SysVar_DirectEmissions_Remelting          = np.einsum('Xn,tn->Xt'    ,RECC_System.ParameterDict['6_PR_DirectEmissions'].Values,SysVar_EnergyDemand_Remelting)
        # Unit: Mt/yr.
        
        # G) Calculate process emissions
        SysVar_ProcessEmissions_PrimaryProd       = np.einsum('mX,tm->Xt'    ,RECC_System.ParameterDict['4_PE_ProcessExtensions_USA'].Values[:,:,110,0],RECC_System.FlowDict['F_3_4'].Values[:,:,0])
        # Unit: Mt/yr.
        
        # H) Calculate emissions from energy supply
        SysVar_IndirectGHGEms_EnergySupply_UsePhase_Buildings      = 0.001 * np.einsum('Xnrt,trgn->Xt',RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply_CDLINKS_MESSAGEix'].Values[:,:,mS,:,:],SysVar_EnergyDemand_UsePhase_Buildings_ByEnergyCarrier)
        SysVar_IndirectGHGEms_EnergySupply_UsePhase_Vehicles       = 0.001 * np.einsum('Xnrt,trgn->Xt',RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply_CDLINKS_MESSAGEix'].Values[:,:,mS,:,:],SysVar_EnergyDemand_UsePhase_Vehicles_ByEnergyCarrier)
        SysVar_IndirectGHGEms_EnergySupply_PrimaryProd             = 0.001 * np.einsum('Xnt,tmn->Xt',  RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply_CDLINKS_MESSAGEix'].Values[:,:,mS,mr,:],SysVar_EnergyDemand_PrimaryProd)
        SysVar_IndirectGHGEms_EnergySupply_Manufacturing           = 0.001 * np.einsum('Xnt,tn->Xt',  RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply_CDLINKS_MESSAGEix'].Values[:,:,mS,mr,:],SysVar_EnergyDemand_Manufacturing)
        SysVar_IndirectGHGEms_EnergySupply_WasteMgt                = 0.001 * np.einsum('Xnt,tn->Xt',  RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply_CDLINKS_MESSAGEix'].Values[:,:,mS,mr,:],SysVar_EnergyDemand_WasteMgt)
        SysVar_IndirectGHGEms_EnergySupply_Remelting               = 0.001 * np.einsum('Xnt,tn->Xt',  RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply_CDLINKS_MESSAGEix'].Values[:,:,mS,mr,:],SysVar_EnergyDemand_Remelting)
        # Unit: Mt/yr.
        
        # I) Calculate total emissions of system
        SysVar_TotalGHGEms = np.einsum('Xtrg->Xt',SysVar_DirectEmissions_UsePhase_Buildings + SysVar_DirectEmissions_UsePhase_Vehicles) + np.einsum('Xtm->Xt',SysVar_DirectEmissions_PrimaryProd) + SysVar_DirectEmissions_Manufacturing + \
        SysVar_DirectEmissions_WasteMgt + SysVar_DirectEmissions_Remelting + SysVar_ProcessEmissions_PrimaryProd + \
        SysVar_IndirectGHGEms_EnergySupply_UsePhase_Buildings + SysVar_IndirectGHGEms_EnergySupply_UsePhase_Vehicles + \
        SysVar_IndirectGHGEms_EnergySupply_PrimaryProd + SysVar_IndirectGHGEms_EnergySupply_Manufacturing +\
        SysVar_IndirectGHGEms_EnergySupply_WasteMgt + SysVar_IndirectGHGEms_EnergySupply_Remelting
        # Unit: Mt/yr.
        
        # I) Calculate indicators
        SysVar_TotalGHGCosts     = np.einsum('t,Xt->Xt',RECC_System.ParameterDict['3_PR_RECC_CO2Price_SSP_32R'].Values[mR,:,mr,mS],SysVar_TotalGHGEms)
        # Unit: million $ / yr.
        
        
#####################################################
#   Section 5) Evaluate results, save, and close    #
#####################################################
Mylog.info('## 5 - Evaluate results, save, and close')
### 5.1.) CREATE PLOTS and include them in log file
Mylog.info('### 5.1 - Create plots and include into logfiles')
Mylog.info('Plot results')

#MyColorCycle = pylab.cm.Paired(np.arange(0,1,0.2))
#linewidth = [1.2,2.4,1.2,1.2,1.2]
#
#Figurecounter = 1
#LegendItems_SSP = IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items
## 1) Emissions by scenario
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#for m in range(0,5):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),SysVar_TotalGHGFootprint[0,:,0,m,5], linewidth = linewidth[m])
#    ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP),shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('GHG emissions of system, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('SSP baseline (unconstrained by RCP)', fontsize = 12) 
#plt.axis([2016, 2050, 0, 1500])
#plt.show()
#fig_name = 'CO2_Ems_Baseline.png'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#for m in range(0,5):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),SysVar_TotalGHGFootprint[0,:,0,m,0], linewidth = linewidth[m])
#    ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP),shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('GHG emissions of system, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('RCP 2.6', fontsize = 12) 
#plt.axis([2016, 2050, 0, 1500])
#plt.show()
#fig_name = 'CO2_Ems_RCP2.6.png'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#for m in range(0,5):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),SysVar_TotalGHGFootprint[0,:,0,m,1], linewidth = linewidth[m])
#    ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP),shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('GHG emissions of system, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('RCP 3.4', fontsize = 12) 
#plt.show()
#fig_name = 'CO2_Ems_RCP3.4.png'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),SysVar_TotalGHGCosts[:,0,:,0]/1000)
#plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('GHG costs of system, bn USD$2005.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.text(2026, (SysVar_TotalGHGCosts[:,0,:,0]/1000).max() * 0.93, '(SSP3 has carbon price of 0.)')
#plt.title('RCP 2.6', fontsize = 12) 
#plt.show()
#fig_name = 'CO2_Ems_Costs_RCP2.6.png'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
##
#
## Primary material production
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),RECC_System.FlowDict['F_3_4'].Values[:,0,:,5,0,:].sum(axis =2))
#plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('Primary production of construction steel, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('SSP baseline (unconstrained by RCP)', fontsize = 12) 
#plt.show()
#fig_name = 'ConstrSteel_Baseline.png'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#for m in range(0,5):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),RECC_System.FlowDict['F_3_4'].Values[:,0,m,5,1,:].sum(axis =1) * 0.15, linewidth = linewidth[m]) # adjusted for cement content
#    ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP),shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('Primary production of automotive steel, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('SSP baseline (unconstrained by RCP)', fontsize = 12) 
#plt.axis([2016, 2050, 0, 4])
#plt.show()
#fig_name = 'AutoSteel_Baseline.png'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#for m in range(0,5):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),RECC_System.FlowDict['F_3_4'].Values[:,0,m,5,12,:].sum(axis =1) * 0.15, linewidth = linewidth[m]) # adjusted for cement content
#    ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP),shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('Production of cement, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('SSP baseline (unconstrained by RCP)', fontsize = 12) 
#plt.axis([2016, 2050, 0, 100])
#plt.show()
#fig_name = 'Cement_Baseline.png'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),RECC_System.FlowDict['F_3_4'].Values[:,0,:,0,0,:].sum(axis =2))
#plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('Primary production of construction steel, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('SSP baseline (RCP 2.6)', fontsize = 12) 
#plt.show()
#fig_name = 'ConstrSteel_RCP2.6.png'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#ProxyHandlesList = []
#for m in range(0,5):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),RECC_System.FlowDict['F_3_4'].Values[:,0,m,0,1,:].sum(axis =1) * 0.15, linewidth = linewidth[m]) # adjusted for cement content
#    ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP),shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('Primary production of automotive steel, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('SSP baseline (RCP 2.6)', fontsize = 12) 
#plt.axis([2016, 2050, 0, 4])
#plt.show()
#fig_name = 'AutoSteel_RCP2.6.png'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#for m in range(0,5):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),RECC_System.FlowDict['F_3_4'].Values[:,0,m,0,12,:].sum(axis =1) * 0.15, linewidth = linewidth[m]) # adjusted for cement content
#    ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP),shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('Production of cement, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('SSP baseline (RCP 2.6)', fontsize = 12) 
#plt.axis([2016, 2050, 0, 100])
#plt.show()
#fig_name = 'Cement_RCP2.6.png'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),RECC_System.FlowDict['F_3_4'].Values[:,0,:,1,0,:].sum(axis =2))
#plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('Primary production of construction steel, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('SSP baseline (RCP 3.4)', fontsize = 12) 
#plt.show()
#fig_name = 'ConstrSteel_RCP3.4.png'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#ProxyHandlesList = []
#for m in range(0,5):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),RECC_System.FlowDict['F_3_4'].Values[:,0,m,1,1,:].sum(axis =1) * 0.15, linewidth = linewidth[m]) # adjusted for cement content
#    ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP),shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('Primary production of automotive steel, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('SSP baseline (RCP 3.4)', fontsize = 12) 
#plt.axis([2016, 2050, 0, 4])
#plt.show()
#fig_name = 'AutoSteel_RCP3.4.png'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#fig1, ax1 = plt.subplots()
#ax1.set_color_cycle(MyColorCycle)
#ProxyHandlesList = []
#for m in range(0,5):
#    ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),RECC_System.FlowDict['F_3_4'].Values[:,0,m,1,12,:].sum(axis =1) * 0.15, linewidth = linewidth[m]) # adjusted for cement content
#    ProxyHandlesList.append(plt.Rectangle((0, 0), 1, 1, fc=MyColorCycle[m,:]))
#plt_lgd  = plt.legend(reversed(ProxyHandlesList),reversed(LegendItems_SSP),shadow = False, prop={'size':9}, loc = 'upper right' ,bbox_to_anchor=(1.20, 1))
#plt.ylabel('Production of cement, Mt/yr.', fontsize = 12) 
#plt.xlabel('year', fontsize = 12) 
#plt.title('SSP baseline (RCP 3.4)', fontsize = 12) 
#plt.axis([2016, 2050, 0, 100])
#plt.show()
#fig_name = 'Cement_RCP3.4.png'
## include figure in logfile:
#fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
#fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500, bbox_inches='tight')
#Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
#Figurecounter += 1
#
#### 5.2) Export to Excel
#Mylog.info('### 5.2 - Export to Excel')
#myfont = xlwt.Font()
#myfont.bold = True
#mystyle = xlwt.XFStyle()
#mystyle.font = myfont
#Result_workbook = xlwt.Workbook(encoding = 'ascii') # Export element stock by region
#
#Sheet = Result_workbook.add_sheet('Cover')
#Sheet.write(2,1,label = 'ScriptConfig', style = mystyle)
#m = 3
#for x in sorted(ScriptConfig.keys()):
#    Sheet.write(m,1,label = x)
#    Sheet.write(m,2,label = ScriptConfig[x])
#    m +=1
#
#MyLabels= []
#for S in range(0,NS):
#    for R in range(0,NR):
#        MyLabels.append(RECC_System.IndexTable.set_index('IndexLetter').loc['S'].Classification.Items[S] + ', ' + RECC_System.IndexTable.set_index('IndexLetter').loc['R'].Classification.Items[R])
#    
#ResultArray = SysVar_TotalGHGFootprint[0,:,0,:,:].reshape(Nt,NS * NR)    
#msf.ExcelSheetFill(Result_workbook, 'TotalGHGFootprint', ResultArray, topcornerlabel = 'System-wide GHG emissions, Mt/yr', rowlabels = RECC_System.IndexTable.set_index('IndexLetter').loc['t'].Classification.Items, collabels = MyLabels, Style = mystyle, rowselect = None, colselect = None)
#
#Result_workbook.save(os.path.join(ProjectSpecs_Path_Result,'SysVar_TotalGHGFootprint.xls'))

### 5.3) Export as .mat file
#Mylog.info('### 5.4 - Export to Matlab')
#Mylog.info('Saving stock data to Matlab.')
#Filestring_Matlab_out = os.path.join(ProjectSpecs_Path_Result, 'StockData.mat')
#scipy.io.savemat(Filestring_Matlab_out, mdict={'Stock': np.einsum('tcrgSe->trgS', RECC_System.StockDict['S_4'].Values)})

### 5.5) Model run is finished. Wrap up.
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
