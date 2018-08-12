# RECC_USA_TestCase_V1.md
Documentation of model building prodedure for US. In a meeting on July 10th, 2018, Edgar, Niko, and Stefan agreed on a 2D-approach to the RECC assessment: a) continue gathering data across the entire system, and b) continue developing the model for the US only.

This piece documents part b) for a first test run of the ODYM-RECC model from scenario drivers to GHG (July 24, 2018).

(1) First, the existing main model script RECC_Main_Development_1.py was copied into RECC_USA_TestCase_V1.py. Then the RECC_Config.xlsx file was opened and a new sheet RECC_USA_TestCase_V1 was created and the content of one of the other config sheets was copied and modified to fit the new model run (change model script, aspect selector, and name of model setting).

(2) A number of dataset files were created and the US-specific data were included.
	+ Refine lifetime dataset 3_LT_RECC_ProductLifetime_V1.0: switch passenger vehicle lifetime to 16 years according to Modaresi and Müller (2012), add all vehicle and building types.
	+ Refine stock dataset 2_S_RECC_FinalProducts_2015_USA_V1.0: Use some crude assumptions to split building stock into energy efficiency classes. 
	  FIX: Vehicle stock data lacking!
	+ Refine material composition datasets 3_MC_RECC_FinalProducts_Vehicles_USA_V1.0 and 3_MC_RECC_FinalProducts_Buildings_USA_V1.0: Use static total values from Hawkins et al. first for vehicles, need to expand to 	components later on. Good variety of sources compiled by Thibaud, that gives enough information to quantify the ca. material content of buildings for a first model version.
	+ Future stock scenario 2_S_RECC_FinalProducts_Future_USA_V1.0: FIX: Made one up ;)
	+ GHG intensity of energy supply 4_PE_GHGIntensityEnergySupply_V1.0: Taken from old IEA model runs, split into BAU (used as proxy for SSPs 3 and 5) and BLUE MAP (used as proxy for SSPs 1, 2, and 4).
	+ EoL recovery rate 4_PY_EoL_RR_USA_V1.0: Taken for Al and Steel from previous work.
	+ RE efficiency strategy potentials: Taken from literature (mostly Allwood et al. estimates), assuptions, and industry estimates for six strategies: more intense use, light-weighting, re-use, product lifetime extension, EoL RR improvement, and fabrication yield improvment.
	+ Intensity of use of products 3_IU_ProductUsePhase: Extrapolate historic data for US. 
	+ Energy intensity of product use 3_EI_ProductUsePhase: Vehicles: Taken from earlier IEA work, via Modaresi et al. (DOI 10.1021/es502930w), electricity for plug-in hybrid vehicles lacking, FIX!. Buildings: FIX: Assumption.

More effort was put into compiling the remaining datasets for the US pilot study: The process yields in the manufacturing and material production sectors and the process extensions in those sectors.
Manufacturing yield loss for steel and aluminium was taken from the literature, the other materials' loss rates were assumed for the first calculation round. 
The remelting yield was compiled from literature and from assumptions.
Process extensions were compiled from ecoinvent.

(3) All dataset files (25 in total) were read by the ODYM dataset parsing routine V2, which reads all datasets into numpy arrays. The MFA system (processes, parameters, flows, stocks) was defined in the ODYM-RECC script RECC_USA_TestCase_V1.0.py, and the model equations for the use phase were programmed as well.
The in-use stock was connected to service provision and to energy consumption, from which the GHG emisisons and the related costs were determined.

Not working as of 2018-07-26: The remelting and fabrication processes have element-specific yield factors, which are not working yet. Need to think more carfully here.

2.8.18:
Add element composition of materials, create loop over model time.

CONVENTION: Stock for year t are measured at the beginning of year t. Model start date is thus 1.1.2016.
