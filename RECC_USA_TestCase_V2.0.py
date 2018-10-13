# -*- coding: utf-8 -*-
"""
Created on July 22, 2018

@authors: spauliuk
"""

"""
File RECC_USA_TestCase_V2.0

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
if Name_Script != 'RECC_USA_TestCase_V2.0':  # Name of this script must equal the specified name in the Excel config file
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
#IndexTable.ix['t']['Classification'].Items # get classification content
Mylog.info('Read model data and parameters.')

ParameterDict = {}
for mo in range(0,len(PL_Names)):
    #ParPath = os.path.join(os.path.abspath(os.path.join(ProjectSpecs_Path_Main, '.')), 'ODYM_RECC_Database', PL_Version[mo])
    ParPath = os.path.join(RECC_Paths.data_path, PL_Version[mo])
    Mylog.info('Reading parameter ' + PL_Names[mo])
    #MetaData, Values = msf.ReadParameter(ParPath = ParPath,ThisPar = PL_Names[mo], ThisParIx = PL_IndexStructure[mo], IndexMatch = PL_IndexMatch[mo], ThisParLayerSel = PL_IndexLayer[mo], MasterClassification,IndexTable,IndexTable_ClassificationNames,ScriptConfig,Mylog) # Do not change order of parameters handed over to function!
    # Do not change order of parameters handed over to function!
    MetaData, Values = msf.ReadParameterV2(ParPath, PL_Names[mo], PL_IndexStructure[mo], PL_IndexMatch[mo],
                                         PL_IndexLayer[mo], MasterClassification, IndexTable,
                                         IndexTable_ClassificationNames, ScriptConfig, Mylog)
    ParameterDict[PL_Names[mo]] = msc.Parameter(Name=MetaData['Dataset_Name'], ID=MetaData['Dataset_ID'],
                                                UUID=MetaData['Dataset_UUID'], P_Res=None, MetaData=MetaData,
                                                Indices=PL_IndexStructure[mo], Values=Values, Uncert=None,
                                                Unit=MetaData['Dataset_Unit'])

# Interpolate missing parameter values:

# 1) Material composition of vehicles:
# Values are given every 5 years, we need all values in between.
index = PL_Names.index('3_MC_RECC_FinalProducts_Vehicles_USA')
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
index = PL_Names.index('3_MC_RECC_FinalProducts_Buildings_USA')
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

# 5) Fabrication yield:
# Extrapolate 2050-2060 as 2050 values
index = PL_Names.index('4_PY_Manufacturing_USA')
ParameterDict[PL_Names[index]].Values[:,:,:,:,1::,:] = np.einsum('t,mwgFr->mwgFtr',np.ones(45),ParameterDict[PL_Names[index]].Values[:,:,:,:,0,:])

# 6 

##########################################################
#    Section 3) Initialize dynamic MFA model for RECC    #
##########################################################
Mylog.info('## 3 - Initialize dynamic MFA model for RECC')
Mylog.info('Define RECC system and processes.')

# Initialize MFA system
RECC_System = msc.MFAsystem(Name='RECC_USA_TestCase_V1.0',
                            Geogr_Scope='R32USA',
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
                                         Indices='t,r,S,R,m,e', Values=None, Uncert=None,
                                         Color=None, ID=None, UUID=None)

RECC_System.FlowDict['F_3_4'] = msc.Flow(Name='primary material production' , P_Start = 3, P_End = 4, 
                                         Indices = 't,r,S,R,m,e', Values=None, Uncert=None, 
                                         Color = None, ID = None, UUID = None)

RECC_System.FlowDict['F_4_5'] = msc.Flow(Name='primary material consumption' , P_Start = 4, P_End = 5, 
                                         Indices = 't,r,S,R,m,e', Values=None, Uncert=None, 
                                         Color = None, ID = None, UUID = None)

RECC_System.FlowDict['F_5_6'] = msc.Flow(Name='manufacturing output' , P_Start = 5, P_End = 6, 
                                         Indices = 't,r,g,S,R,m,e', Values=None, Uncert=None, 
                                         Color = None, ID = None, UUID = None)
    
RECC_System.FlowDict['F_6_7'] = msc.Flow(Name='Final consumption', P_Start=6, P_End=7,
                                         Indices='t,r,g,S,R,m,e', Values=None, Uncert=None,
                                         Color=None, ID=None, UUID=None)

RECC_System.FlowDict['F_7_8'] = msc.Flow(Name='EoL products' , P_Start = 7, P_End = 8, 
                                         Indices = 't,c,r,g,S,R,m,e', Values=None, Uncert=None, 
                                         Color = None, ID = None, UUID = None)

RECC_System.FlowDict['F_8_9'] = msc.Flow(Name='Waste mgt. input' , P_Start = 8, P_End = 9, 
                                         Indices = 't,r,g,S,R,m,e', Values=None, Uncert=None, 
                                         Color = None, ID = None, UUID = None)

RECC_System.FlowDict['F_8_6'] = msc.Flow(Name='Product re-use' , P_Start = 8, P_End = 6, 
                                         Indices = 't,r,g,S,R,m,e', Values=None, Uncert=None, 
                                         Color = None, ID = None, UUID = None)

RECC_System.FlowDict['F_9_10'] = msc.Flow(Name='old scrap' , P_Start = 9, P_End = 10, 
                                         Indices = 't,r,S,R,w,e', Values=None, Uncert=None, 
                                         Color = None, ID = None, UUID = None)

RECC_System.FlowDict['F_5_10'] = msc.Flow(Name='New scrap' , P_Start = 5, P_End = 10, 
                                         Indices = 't,r,S,R,w,e', Values=None, Uncert=None, 
                                         Color = None, ID = None, UUID = None)

RECC_System.FlowDict['F_10_11'] = msc.Flow(Name='scrap use' , P_Start = 10, P_End = 11, 
                                         Indices = 't,r,S,R,w,e', Values=None, Uncert=None, 
                                         Color = None, ID = None, UUID = None)

RECC_System.FlowDict['F_11_12'] = msc.Flow(Name='secondary material production' , P_Start = 11, P_End = 12, 
                                         Indices = 't,r,S,R,m,e', Values=None, Uncert=None, 
                                         Color = None, ID = None, UUID = None)

RECC_System.FlowDict['F_12_5'] = msc.Flow(Name='secondary material consumption' , P_Start = 12, P_End = 5, 
                                         Indices = 't,r,S,R,m,e', Values=None, Uncert=None, 
                                         Color = None, ID = None, UUID = None)

RECC_System.FlowDict['F_9_0'] = msc.Flow(Name='waste mgt. losses' , P_Start = 9, P_End = 0, 
                                         Indices = 't,r,S,R,e', Values=None, Uncert=None, 
                                         Color = None, ID = None, UUID = None)

RECC_System.FlowDict['F_11_0'] = msc.Flow(Name='remelting. losses' , P_Start = 11, P_End = 0, 
                                         Indices = 't,r,S,R,e', Values=None, Uncert=None, 
                                         Color = None, ID = None, UUID = None)

# Define system variables: Stocks.
RECC_System.StockDict['S_7']   = msc.Stock(Name='In-use stock', P_Res=7, Type=1,
                                         Indices = 't,c,r,g,S,R,m,e', Values=None, Uncert=None,
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
#TotalStockCurves_UsePhase   = np.zeros((Nt,Nr,NG,NS))    # Stock   by year, region, scenario, and product
#TotalStockCurves_C    = np.zeros((Nt,Nc,Nr,NG,NS)) # Stock   by year, age-cohort, region, scenario, and product
#TotalInflowCurves     = np.zeros((Nt,Nr,NG,NS))    # Inflow  by year, region, scenario, and product
#TotalOutflowCurves    = np.zeros((Nt,Nc,Nr,NG,NS)) # Outflow by year, age-cohort, region, scenario, and product
SF_Array               = np.zeros((Nc,Nc,Ng,Nr,NS,NR)) # survival functions, by year, age-cohort, good, region, SSP scenario, and RCP scenario. PDFs are stored externally because recreating them with scipy.stats is slow.
Stock_Detail_UsePhase      = np.zeros((Nt,Nc,Ng,Nr,NS,NR)) # index structure: tcgrSR
Outflow_Detail_UsePhase     = np.zeros((Nt,Nc,Ng,Nr,NS,NR)) # index structure: tcgrSR
Inflow_Detail_UsePhase      = np.zeros((Nt,Ng,Nr,NS,NR)) # index structure: tgrSR
SwitchTime=Nc - Model_Duration +1 # Year when future modelling horizon starts: 1.1.2016

#Get historic stock in 2015 by age-cohort, and covert unit to Vehicles: million, Buildings: million m2.
TotalStock_UsePhase_Hist_cgr = RECC_System.ParameterDict['2_S_RECC_FinalProducts_2015_USA'].Values[0,:,:,:].copy() 

# Determine total future stock, product level. Units: Vehicles: million, Buildings: million m2.
TotalStockCurves_UsePhase = np.einsum('StGr,MtrS->trGS',RECC_System.ParameterDict['2_S_RECC_FinalProducts_Future_USA'].Values,RECC_System.ParameterDict['2_P_RECC_Population_SSP_32R'].Values)

# 2) Include (or not) the RE strategies for the use phase:

# Include_REStrategy_MoreIntenseUse:
if ScriptConfig['Include_REStrategy_MoreIntenseUse'] == 'True':
    TotalStockCurves_UsePhase = np.einsum('RtrgS,trgS->RtrgS', (1 - np.einsum('RtrS,grS->RtrgS',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp_SSP_32R'].Values*0.01,RECC_System.ParameterDict['6_PR_MoreIntenseUse_USA'].Values)),TotalStockCurves_UsePhase)    
else:
    TotalStockCurves_UsePhase = np.einsum('R,trgS->RtrgS', np.ones(NR),TotalStockCurves_UsePhase)    

# Include_REStrategy_LifeTimeExtension: Product lifetime extension.
# First, replicate lifetimes for the 5 scenarios and all age-cohorts
Par_RECC_ProductLifetime = np.einsum('ScR,gUr->RgrcS',np.ones((NS,Nc,NR)),RECC_System.ParameterDict['3_LT_RECC_ProductLifetime'].Values)
# Second, change lifetime of future age-cohorts according to lifetime extension parameter
# This is equation 10 of the paper:
if ScriptConfig['Include_REStrategy_LifeTimeExtension'] == 'True':
    Par_RECC_ProductLifetime[:,:,:,SwitchTime -1::,:] = np.einsum('RcrgS,RgrcS->RgrcS',1 + np.einsum('RcrS,grS->RcrgS',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp_SSP_32R'].Values*0.01,RECC_System.ParameterDict['6_PR_LifeTimeExtension_USA'].Values),Par_RECC_ProductLifetime[:,:,:,SwitchTime -1::,:])

# 3) Dynamic stock model
# Build pdf array from lifetime distribution: Probability of survival.
for g in tqdm(range(0, Ng), unit=' commodity groups'):
    for r in range(0, Nr):
        for S in tqdm(range(0,NS), unit='SSP scenarios'):
            for R in range(0,NR):
                LifeTimes = Par_RECC_ProductLifetime[R, g, r, :, S]
                lt = {'Type'  : 'Normal',
                      'Mean'  : LifeTimes,
                      'StdDev': 0.3 * LifeTimes}
                SF_Array[:, :, g, r, S, R] = dsm.DynamicStockModel(t=np.arange(0, Nc, 1), lt=lt).compute_sf().copy()
                # np.fill_diagonal(SF_Array[:, :, g, r, S, R],1) # no outflows from current year, this would break the mass balance in the calculation routine below.

# Compute evolution of 2015 in-use stocks: initial stock evolution separately from future stock demand and stock-driven model
for G in tqdm(range(0, NG), unit=' commodity groups'):
    for r in range(0,Nr):
        for S in tqdm(range(0,NS), unit='SSP scenarios'):
            for R in range(0,NR):    
                FutureStock     = TotalStockCurves_UsePhase[R,1::, r, G, S]# Future total stock
                InitialStock    = TotalStock_UsePhase_Hist_cgr[:,:,r]
                if G == 0: # quick fix for vehicles and buildings only !!!
                    InitialStock[:,6::] = 0  # set not relevant initial stock to 0
                if G == 1:
                    InitialStock[:,0:6] = 0  # set not relevant initial stock to 0
                SFArrayCombined = SF_Array[:,:,:,r,S,R]
                TypeSplit       = RECC_System.ParameterDict['3_SHA_TypeSplit_NewProducts_USA'].Values[G,:,1::,r,S].transpose()
  
                Var_S, Var_O, Var_I = msf.compute_stock_driven_model_initialstock_typesplit(FutureStock,InitialStock,SFArrayCombined,TypeSplit, NegativeInflowCorrect = False)
                
                Stock_Detail_UsePhase[1::,:,:,r,S,R]   += Var_S.copy()
                Outflow_Detail_UsePhase[1::,:,:,r,S,R] += Var_O.copy()
                Inflow_Detail_UsePhase[1::,:,r,S,R]    += Var_I[SwitchTime::,:].copy()
# Here so far: Units: Vehicles: million, Buildings: million m2. for stocks, X/yr for flows.
                
# Clean up
del TotalStockCurves_UsePhase
del SF_Array

# include light-weighting:
Par_RECC_MC_SR = np.einsum('cmgr,SR->cmgrSR',RECC_System.ParameterDict['3_MC_RECC_FinalProducts_Buildings_USA'].Values[:,1,:,:,:] + RECC_System.ParameterDict['3_MC_RECC_FinalProducts_Vehicles_USA'].Values[:,0,:,:,:],np.ones((NS,NR)))
if ScriptConfig['Include_REStrategy_LightWeighting'] == 'True':
    Par_RECC_MC_SR[SwitchTime-1::,:,:,:,:,:] = np.einsum('cmgrSR,RcrgS->cmgrSR',Par_RECC_MC_SR[SwitchTime-1::,:,:,:,:,:], 1 - np.einsum('RcrS,grS->RcrgS',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp_SSP_32R'].Values*0.01,RECC_System.ParameterDict['6_PR_LightWeighting_USA'].Values))
# Units: Vehicles: kg/unit, Buildings: kg/m2  

# 4) Build dynamic MFA system
# Calculate material stocks and flows, use phase. Units: Mt and Mt/yr.
# This calculation is done year-by-year, and the elemental composition of the materials is in part determined by the scrap flow metal composition
Par_Element_Composition_of_Materials = np.zeros((Nc,Nr,Ng,Nm,Ne,NS,NR)) # crgmeSR
Par_Element_Composition_of_Materials[0:Nc-Nt+1,:,:,:,:,:,:] = np.einsum('crgSR,me->crgmeSR',np.ones((Nc -Nt +1,Nr,Ng,NS,NR)),RECC_System.ParameterDict['3_MC_Elements_Materials_ExistingStock'].Values)

# know total inflows and outflows by material, need to split them into elements later:
RECC_System.FlowDict['F_6_7'].Values[:,:,:,:,:,:,0]   = \
np.einsum('tmgrSR,tgrSR->trgSRm',Par_RECC_MC_SR[SwitchTime-1::,:,:,:,:,:],Inflow_Detail_UsePhase)/1000 # Indices='t,r,g,S,R,m,e'

# Memory-intensive:
RECC_System.FlowDict['F_7_8'].Values[:,:,:,:,:,:,:,0] = \
np.einsum('cmgrSR,tcgrSR->tcrgSRm',Par_RECC_MC_SR,Outflow_Detail_UsePhase)/1000
# 1) expand chemical elements of outflow:
RECC_System.FlowDict['F_7_8'].Values[0:2,0:Nc-Nt+1,:,:,:,:,:,1::] = np.einsum('tcrgSRm,crgmeSR->tcrgSRme',RECC_System.FlowDict['F_7_8'].Values[0:2,0:Nc-Nt+1,:,:,:,:,:,0],Par_Element_Composition_of_Materials[0:Nc-Nt+1,:,:,:,1::,:,:])

# Memory-intensive:
RECC_System.StockDict['S_7'].Values[:,:,:,:,:,:,:,0]   = \
np.einsum('cmgrSR,tcgrSR->tcrgSRm',Par_RECC_MC_SR,Stock_Detail_UsePhase)/1000

# Prepare parameters:
# Manufacturing yield:
Par_FabYield = np.einsum('mwggtr,SR->mwgtrSR',RECC_System.ParameterDict['4_PY_Manufacturing_USA'].Values,np.ones((NS,NR))) # take diagonal of product = manufacturing process
# Consider Fabrication yield improvement
if ScriptConfig['Include_REStrategy_FabYieldImprovement'] == 'True':
    Par_FabYieldImprovement = np.einsum('w,RtmgrS->mwgtrSR',np.ones((Nw)),np.einsum('RtrS,mgrS->RtmgrS',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp_SSP_32R'].Values*0.01,RECC_System.ParameterDict['6_PR_FabricationYieldImprovement_USA'].Values))
else:
    Par_FabYieldImprovement = 0
Par_FabYield_Raster = Par_FabYield > 0    
Par_FabYield        = Par_FabYield - Par_FabYield_Raster * Par_FabYieldImprovement
Par_FabYield_total  = np.einsum('mwgtrSR->mgtrSR',Par_FabYield)

# Consider EoL recovery rate improvement:
if ScriptConfig['Include_REStrategy_EoL_RR_Improvement'] == 'True':
    Par_RECC_EoL_RR_USA_SR =  np.einsum('RtS,grmw->RtrmgSw',np.ones((NR,Nt,NS)),RECC_System.ParameterDict['4_PY_EoL_RR_USA'].Values *0.01) \
    + np.einsum('RtrS,grmw->RtrmgSw',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp_SSP_32R'].Values*0.01,RECC_System.ParameterDict['6_PR_EoL_RR_Improvement_USA'].Values*0.01)
else:    
    Par_RECC_EoL_RR_USA_SR =  np.einsum('RtS,grmw->RtrmgSw',np.ones((NR,Nt,NS)),RECC_System.ParameterDict['4_PY_EoL_RR_USA'].Values *0.01)

# Continue to develop system year by year to transfer total material flows to element level
for t in tqdm(range(1, Nt), unit=' years'): # 1: 2016
    CohortOffset = t +Nc -Nt    
    
    # waste mgt and re-use:
    # 1) expand chemical elements of outflow:
    RECC_System.FlowDict['F_7_8'].Values[t,0:CohortOffset,:,:,:,:,:,1::] = np.einsum('crgSRm,crgmeSR->crgSRme',RECC_System.FlowDict['F_7_8'].Values[t,0:CohortOffset,:,:,:,:,:,0],Par_Element_Composition_of_Materials[0:CohortOffset,:,:,:,1::,:,:])

    # Consider Re-use
    if ScriptConfig['Include_REStrategy_ReUse'] == 'True':
        RECC_System.FlowDict['F_8_6'].Values[t,:,:,:,:,:,:] =  np.einsum('RrgS,crgSRme->rgSRme',np.einsum('RrS,grS->RrgS',RECC_System.ParameterDict['3_SHA_RECC_REStrategyScaleUp_SSP_32R'].Values[:,t,:,:]*0.01,RECC_System.ParameterDict['6_PR_ReUse_USA'].Values),RECC_System.FlowDict['F_7_8'].Values[t,0:Nc-Nt+t,:,:,:,:,:,:])
        
    RECC_System.FlowDict['F_8_9'].Values[t,:,:,:,:,:,:]     = np.einsum('crgSRme->rgSRme',RECC_System.FlowDict['F_7_8'].Values[t,:,:,:,:,:,:]) - RECC_System.FlowDict['F_8_6'].Values[t,:,:,:,:,:,:]
    
    # Calculate postconsumper scrap, in Mt/yr
    RECC_System.FlowDict['F_9_10'].Values[t,:,:,:,:,:]      = np.einsum('RrmgSw,rgSRme->rSRwe',Par_RECC_EoL_RR_USA_SR[:,t,:,:,:,:,:],RECC_System.FlowDict['F_8_9'].Values[t,:,:,:,:,:,:])    
        
    # waste mgt. losses:
    RECC_System.FlowDict['F_9_0'].Values[t,:,:,:,:]         = np.einsum('rgSRme->rSRe',RECC_System.FlowDict['F_8_9'].Values[t,:,:,:,:,:,:]) - np.einsum('rSRwe->rSRe',RECC_System.FlowDict['F_9_10'].Values[t,:,:,:,:,:])
    
    # remelting. 
    # Add old scrap with manufacturing scrap from last year. In year 2016, no fabrication scrap exists yet.
    RECC_System.FlowDict['F_10_11'].Values[t,:,:,:,:,:] = RECC_System.FlowDict['F_9_10'].Values[t,:,:,:,:,:] + RECC_System.FlowDict['F_5_10'].Values[t-1,:,:,:,:,:]
    RECC_System.FlowDict['F_11_12'].Values[t,:,:,:,:,:] = np.einsum('rSRwe,wmePr->rSRme',RECC_System.FlowDict['F_10_11'].Values[t,:,:,:,:,:],RECC_System.ParameterDict['4_PY_MaterialProductionRemelting'].Values[:,:,:,:,t,:])
    RECC_System.FlowDict['F_11_0'].Values[t,:,:,:,:]    = np.einsum('rSRwe->rSRe',RECC_System.FlowDict['F_10_11'].Values[t,:,:,:,:,:]) - np.einsum('rSRwe->rSRe',RECC_System.FlowDict['F_11_12'].Values[t,:,:,:,:,:])
    
    RECC_System.FlowDict['F_12_5'].Values[t,:,:,:,:,:]  = RECC_System.FlowDict['F_11_12'].Values[t,:,:,:,:,:]

    # Calculate manufacturing output, in Mt/yr
    RECC_System.FlowDict['F_5_6'].Values[t,:,:,:,:,:,:] = RECC_System.FlowDict['F_6_7'].Values[t,:,:,:,:,:,:] - RECC_System.FlowDict['F_8_6'].Values[t,:,:,:,:,:,:]

    # Calculate manufacturing input and fabrication scrap
    Par_FabYield_total_inv = 1/(1-Par_FabYield_total)
    Par_FabYield_total_inv[Par_FabYield_total_inv == 1] = 0
    RECC_Flow_Manufacturing_Input_Total_t = np.einsum('mgrSR,rgSRme->rSRme',Par_FabYield_total_inv[:,:,t,:,:,:],RECC_System.FlowDict['F_5_6'].Values[t,:,:,:,:,:,:])

    RECC_System.FlowDict['F_4_5'].Values[t,:,:,:,:,:] = RECC_Flow_Manufacturing_Input_Total_t - RECC_System.FlowDict['F_12_5'].Values[t,:,:,:,:,:]
    RECC_System.FlowDict['F_3_4'].Values[t,:,:,:,:,:] = RECC_System.FlowDict['F_4_5'].Values[t,:,:,:,:,:]

    #Fabrication scrap, to be recycled next year:
    RECC_System.FlowDict['F_5_10'].Values[t,:,:,:,:,:] = \
    np.einsum('mwgrSR,rSRme->rSRwe',Par_FabYield[:,:,:,t,:,:,:],\
    (RECC_System.FlowDict['F_4_5'].Values[t,:,:,:,:,:] + RECC_System.FlowDict['F_4_5'].Values[t,:,:,:,:,:]))

    # Calculate net ore input to primary production, mining and refining will be described in detail later:
    RECC_System.FlowDict['F_0_3'].Values[t,:,:,:,:,:] = RECC_System.FlowDict['F_3_4'].Values[t,:,:,:,:,:]
    
    # Calculate element composition of material and inflow:
    RECC_System.FlowDict['F_6_7'].Values[t,:,:,:,:,:,1::] = RECC_System.FlowDict['F_5_6'].Values[t,:,:,:,:,:,1::] + RECC_System.FlowDict['F_8_6'].Values[t,:,:,:,:,:,1::]
    ThisYear_MatComposition_inv = 1/ RECC_System.FlowDict['F_5_6'].Values[t,:,:,:,:,:,1::].sum(axis =5)
    ThisYear_MatComposition_inv[np.isinf(ThisYear_MatComposition_inv)] = 0
    ThisYear_MatComposition = np.einsum('rgSRme,rgSRm->rgSRme',RECC_System.FlowDict['F_5_6'].Values[t,:,:,:,:,:,1::],ThisYear_MatComposition_inv)
    Par_Element_Composition_of_Materials[CohortOffset,:,:,:,1::,:,:] = np.einsum('rgSRme->rgmeSR',ThisYear_MatComposition)

# Calculate element composition of stock:
RECC_System.StockDict['S_7'].Values[:,:,:,:,:,:,:,1::] = np.einsum('tcrgSRm,crgmeSR->tcrgSRme',RECC_System.StockDict['S_7'].Values[:,:,:,:,:,:,:,0],Par_Element_Composition_of_Materials[:,:,:,:,1::,:,:])
#Units so far: Mt/yr

# Determine energy demand, environmental extensions, and CO2 costs
SysVar_ServiceSupply_UsePhase = np.einsum('tgrS,tcgrSR->tcgrSR',RECC_System.ParameterDict['3_IU_Products_UsePhase_USA'].Values, Stock_Detail_UsePhase)
# Unit: million km/yr for vehicles, million m2 for buildings

SysVar_EnergyDemand_UsePhase  = np.einsum('cgnrS,tcgrSR->trnSR',RECC_System.ParameterDict['3_EI_Products_UsePhase_USA'].Values, SysVar_ServiceSupply_UsePhase)
# Unit: TJ/yr for both vehicles and buildings.

SysVar_EnergyDemand_MaterialRemelting = 1000 * np.einsum('Pnr,trSRme->trnSR',RECC_System.ParameterDict['4_EI_ProcessEnergyIntensity_USA'].Values[0:7,:,110,:],RECC_System.FlowDict['F_11_12'].Values)
# Unit: TJ/yr.

SysVar_TotalEnergyDemand = SysVar_EnergyDemand_UsePhase + SysVar_EnergyDemand_MaterialRemelting
# Unit: TJ/yr.

SysVar_DirectGHGEms_UsePhase            = 0.001 * np.einsum('Xn,trnSR->XtrSR',RECC_System.ParameterDict['6_PR_DirectEmissions'].Values,SysVar_TotalEnergyDemand)
# Unit: Mt/yr.

SysVar_TotelEnergyGHGFootprint          = 0.001 * np.einsum('XnSt,trnSR->XtrSR',RECC_System.ParameterDict['4_PE_GHGIntensityEnergySupply'].Values,SysVar_TotalEnergyDemand)
# Unit: Mt/yr.

#SysVar_TotalPrimaryMaterialGHGFootprint = np.einsum('mXr,trSRm->XtrSR',RECC_System.ParameterDict['4_PE_ProcessExtensions_USA'].Values[7::,:,110,:],SysVar_PrimaryProd[:,:,:,:,:,0])
SysVar_TotalPrimaryMaterialGHGFootprint = np.einsum('mXr,trSRme->XtrSR',RECC_System.ParameterDict['4_PE_ProcessExtensions_USA'].Values[7::,:,110,:],RECC_System.FlowDict['F_3_4'].Values)
# Unit: Mt/yr.  Here, we have a one to one correspondence between materials and productions processes, hence, P = m.

SysVar_TotalGHGFootprint                = SysVar_TotelEnergyGHGFootprint + SysVar_DirectGHGEms_UsePhase + SysVar_TotalPrimaryMaterialGHGFootprint
# Dimensions: extension x time x region x SSP x RCP
# Unit: Mt/yr.

SysVar_TotalGHGCosts     = np.einsum('RtrS,XtrSR->trSR',RECC_System.ParameterDict['3_PR_RECC_CO2Price_SSP_32R'].Values,SysVar_TotalGHGFootprint)
# Unit: million $ / yr.


RECC_System.Consistency_Check() # Check whether flow value arrays match their indices, etc.

# Determine Mass Balance
Bal = RECC_System.MassBalance()

#####################################################
#   Section 5) Evaluate results, save, and close    #
#####################################################
Mylog.info('## 5 - Evaluate results, save, and close')
### 5.1.) CREATE PLOTS and include them in log file
Mylog.info('### 5.1 - Create plots and include into logfiles')
Mylog.info('Plot results')
Figurecounter = 1
LegendItems_SSP = IndexTable.Classification[IndexTable.index.get_loc('Scenario')].Items
# 1) Emissions by scenerio

fig1, ax1 = plt.subplots()
ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),SysVar_TotalGHGFootprint[0,:,0,:,5])
plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9},loc='upper left')
plt.ylabel('GHG emissions of system, Mt/yr.', fontsize = 12) 
plt.xlabel('year', fontsize = 12) 
plt.title('SSP baseline (unconstrained by RCP)', fontsize = 12) 
plt.show()
fig_name = 'CO2_Ems_Baseline.png'
# include figure in logfile:
fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500)
Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
Figurecounter += 1

fig1, ax1 = plt.subplots()
ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),SysVar_TotalGHGFootprint[0,:,0,:,0])
plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9},loc='upper left')
plt.ylabel('GHG emissions of system, Mt/yr.', fontsize = 12) 
plt.xlabel('year', fontsize = 12) 
plt.title('RCP 2.6', fontsize = 12) 
plt.show()
fig_name = 'CO2_Ems_RCP2.6.png'
# include figure in logfile:
fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500)
Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
Figurecounter += 1

fig1, ax1 = plt.subplots()
ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),SysVar_TotalGHGFootprint[0,:,0,:,1])
plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9},loc='upper left')
plt.ylabel('GHG emissions of system, Mt/yr.', fontsize = 12) 
plt.xlabel('year', fontsize = 12) 
plt.title('RCP 3.4', fontsize = 12) 
plt.show()
fig_name = 'CO2_Ems_RCP3.4.png'
# include figure in logfile:
fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500)
Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
Figurecounter += 1

fig1, ax1 = plt.subplots()
ax1.plot(np.arange(Model_Time_Start,Model_Time_End +1),SysVar_TotalGHGCosts[:,0,:,0]/1000)
plt_lgd  = plt.legend(LegendItems_SSP,shadow = False, prop={'size':9},loc='upper left')
plt.ylabel('GHG costs of system, bn USD$2005.', fontsize = 12) 
plt.xlabel('year', fontsize = 12) 
plt.text(2026, 1050, '(SSP3 has carbon price of 0.)')
plt.title('RCP 2.6', fontsize = 12) 
plt.show()
fig_name = 'CO2_Ems_Costs_RCP2.6.png'
# include figure in logfile:
fig_name = 'Figure ' + str(Figurecounter) + '_' + fig_name
fig1.savefig(os.path.join(ProjectSpecs_Path_Result, fig_name), dpi=500)
Mylog.info('![%s](%s){ width=850px }' % (fig_name, fig_name))
Figurecounter += 1
#

### 5.2) Export to Excel
Mylog.info('### 5.2 - Export to Excel')
myfont = xlwt.Font()
myfont.bold = True
mystyle = xlwt.XFStyle()
mystyle.font = myfont
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
    
ResultArray = SysVar_TotalGHGFootprint[0,:,0,:,:].reshape(Nt,NS * NR)    
msf.ExcelSheetFill(Result_workbook, 'TotalGHGFootprint', ResultArray, topcornerlabel = 'System-wide GHG emissions, Mt/yr', rowlabels = RECC_System.IndexTable.set_index('IndexLetter').loc['t'].Classification.Items, collabels = MyLabels, Style = mystyle, rowselect = None, colselect = None)

Result_workbook.save(os.path.join(ProjectSpecs_Path_Result,'SysVar_TotalGHGFootprint.xls'))

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
