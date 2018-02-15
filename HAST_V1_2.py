"""
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
NAME:             HAST Harmonics Automated Simulation Tool
VERSION:          1.2 [24 April 2017]
AUTHOR:           Barry O'Connell
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

David Mills [DM] from PSC UK rewrote this to improve performance and avoid some errors that seemed to be in the code
	Code layout is now as per PEP

-------------------------------------------------------------------------------------------------------------------
IMPORTANT NOTES:

Code uses TABS for indenting rather than SPACES.

Notepad ++ is a useful tool for viewing python coded (DM - recommends PyCharm instead as it conforms to PEP)
Install Python 3.5 to your C:\ or D:\ drive. Do not install in C:\\programfiles  as the win32com module needs write
	cess to create a cache and it wont have that in program files
Use these commands to check Environment variables are setup correctly
1. help("modules")
	1.1 Check if powerfactory is in your modules, if not copy powerfactory python dll
		(c:\\programfiles\\pf etc) to python directory (eg C:\\python3.5\\DLL)
2. print(os.environ["PATH"])
	2.1 Check to see the correct path above was appended to your environment variables

3. for param in os.environ.keys():
	print "%20s %s" % (param,os.environ[param])

4. If you are having trouble with numpy scipy ensure that you either install the modules or anaconda
	which has these modules present
5. You can comment out numpy and scipy in the import section if you set Excel_Convex_Hull = False. This will
	then skip creating the points for the convex hull

---------------------------------------------------------------------------------------------------------------------
UNIT TESTING
Unit tests have begun to be added to this script.  When any changes are made it is recommended to run the unittests
to determine that the code in principal works correctly.

----------------------------------
SIGNIFICANT UPDATES
- Converted to us a logging system that is stored in hast.logger which will avoid writing to a log file every time
something happens

- Working on:
	1. Using win32com to write to named chart rather than active chart to speed up chart performance


"""

# IMPORT SOME PYTHON MODULES
# --------------------------------------------------------------------------------------------------------------------
import os
import sys
import importlib

import powerfactory 					# Power factory module see notes above
import time                          	# Time
import ctypes                        	# For creating startup message box
import re								# Used for stripping text strings
import unittest							# Used to include test functions for error checking of code

# HAST module package requires reload during code development since python does not reload itself
# HAST module package used as functions start to be transferred for efficiency
import hast
hast = importlib.reload(hast)

# GLOBAL variable used to avoid trying to print to PowerFactory when running in unittest mode, set to true by unittest
DEBUG_MODE = False

# Functions -----------------------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------------------------------
def print1(name, bf=0, af=0):   # Used to print a message to both python, PF and write it to file with double space
	"""
		Used to print a message to both python, PF and write it to file with double space
	:param int bf: Number of blank lines before statement in Progress file
	:param str name: Message to display
	:param int af: Number of blank lines after statement in Progress file
	:return: None
	"""

	# Updated to now use logging to control printout
	# First call to logger occurs before it has been declared and therefore if this is the case a simple print of the
	# message is performed.  This should only happen because a status message is provided when the excel instance is
	# initiated and this occurs before the excel workbook containing details of the log files has been imported.
	# TODO:  Future update to logger to use debug log and change the target file after reading in the workbook
	try:
		logger.info(name)
	except NameError:
		print(name)
	return None


def print2(name, bf=2, af=0):   # Used to print error message to both python, PF and write it to the file
	"""
		Used to print error message to both python, PF and write it to the file
	:param str name:  Error message to display
	:param int bf: (optional) = Number of empty lines before error message
	:param int af: (optional) = Number of empty lines after error message
	:return: None
	"""
	# Updated to use logging handler for error messages
	# First call to logger occurs before it has been declared and therefore if this is the case a simple print of the
	# message is performed.  This should only happen because a status message is provided when the excel instance is
	# initiated and this occurs before the excel workbook containing details of the log files has been imported.
	# TODO:  Future update to logger to use debug log and change the target file after reading in the workbook
	try:
		logger.error(name)
	except NameError:
		print(name)
	return None


def activate_project(project): 		# Activate project
	"""
		Activate project in Power Factory
	:param str project: Name of project to be activated 
	:return powerfactory.Project _prj: Handle for powerfactory project 
	"""
	pro = app.ActivateProject(project) 										# Activate project
	if pro == 0:															# Project Activate Successfully
		# Print Information to progress log and PowerFactory window
		print1('Activated Project Successfully: {}'.format(project), bf=1, af=0)
		# prj renamed _prj to avoid shadowing name from parent project
		_prj = app.GetActiveProject()										# Get active project
	else:																	# Project Failed to Activate
		# Print Information to progress log and PowerFactory window and Error Log
		print2(('Error Not able to Activate Project: {}.........................'.format(project)))
		# prj renamed _prj to avoid shadowing name from parent project
		_prj = []
	return _prj


def activate_study_case(study_case): 		# Activate Study case
	"""
		Activate Study case
	:param str study_case: Study case to be checked 
	:return (list int) (study_case[0], cas): Handle for activated study case, 0 if successful 
	"""
	deactivate_study_case()
	study_case_folder1 = app.GetProjectFolder("study")							# Returns string the location of the project folder for study cases, scen,
	study_case1 = study_case_folder1.GetContents(study_case)
	if len(study_case1) > 0:
		cas = study_case1[0].Activate() 														# Activate Study case
		if cas == 0:
			print1('Activated Study Case Successfully: {}'.format(study_case1[0]), bf=1, af=0)
		else:
			print2('Error Unsuccessfully Activated Study Case: {}.............................'.format(study_case))
	else:
		print2('Could not activate StudyCase as no matching name in case: {}'.format(study_case))
		cas = 1
		study_case1 = [[]]

	# Returns handle for study_case and identifier of 0 if case load is successful
	return study_case1[0], cas


def deactivate_study_case(): 		# Deactivate Scenario
	"""
		Deactivate loaded study case
	:return: None
	"""
	# Get handle for active study case from PowerFactory
	study = app.GetActiveStudyCase()
	if study is not None:
		sce = study.Deactivate() 											# Deactivate Study case
		if sce == 0:
			pass
			# TODO: Should add a debug message here
			# print1(1,"Deactivated Active Study Case Successfully : " + str(Study),0)
		elif sce > 0:
			print2('Error Unsuccessfully Deactivated Study Case: {}..............................'.format(study))
			print2('Unsuccessfully Deactivated Scenario Error Code: {}'.format(sce))
	else:
		print1("No Study Case Active to Deactivate ................................", bf=2, af=0)
	return None


def activate_scenario(scenario): 		# Activate Scenario
	"""
		Activate scenario
	:param str scenario: Name of scenario to activate 
	:return (scenario1[0], sce:  
	"""
	scenario_folder1 = app.GetProjectFolder("scen")							# Returns string the location of the project folder for study cases, scen,
	scenario1 = scenario_folder1.GetContents(scenario)
	deactivate_scenario()
	#print2('Scenarios :{}'.format(scenario1))
	sce = scenario1[0].Activate() 											# Activate Study case
	if sce == 0:
		print1('Activated Scenario Successfully: {}'.format(scenario1[0]), bf=1, af=0)
	elif sce > 0:
		print2('Error Unsuccessfully Activated Scenario: {}.........................'.format(scenario1[0]))
		print2('Unsuccessfully Activated Scenario Error Code: {}'.format(sce))
	return scenario1[0], sce


def activate_scenario1(scenario): 		# Activate Scenario
	"""
		Activate scenario
	:param powerfactory.Scenario scenario: Activates scenario passed as input handle 
	:return: status on attempt to activate
	"""
	sce = scenario.Activate() 											# Activate Study case
	if sce == 0:
		print1('Activated Scenario Successfully: {}'.format(scenario), bf=1, af=0)
	elif sce == 1:
		print2('Error Unsuccessfully Activated Scenario: {}...............................'.format(scenario))
		print2('Unsuccessfully Activated Scenario Error Code: {}'.format(sce))
	return sce


def deactivate_scenario(): 		# Deactivate Scenario
	"""
		Deactivate the active scenario
	:return: None
	"""
	scenario1 = app.GetActiveScenario()
	# Only deactivate a scenario if it already exists
	if scenario1 is not None:
		sce = scenario1.Deactivate() 											# Deactivate Study case
		if sce == 0:
			pass
			# TODO:  Should add in debug statement if successful
			# print1(1,("Deactivated Active Scenario Successfully : " + str(Scenario1)),0)
		elif sce > 0:
			print2('Error Unsuccessfully Deactivated Scenario: {}..............................'.format(scenario1))
			print2('Unsuccessfully Deactivated Scenario Error Code: {}'.format(sce))
	else:
		print1('No Scenario Active to Deactivate ................................', bf=1, af=0)
	return None


def save_active_scenario(): 		# Save active scenario
	"""
		Save the active scenario
	:return: None
	"""
	scenario1 = app.GetActiveScenario()
	sce = scenario1.Save()
	if sce==0:
		print1('Saved active scenario successfully: {}'.format(scenario1), bf=1, af=0)
	elif sce == 1 and scenario1 is None:
		print2('Error unsuccessfully saved scenario: {}'.format(scenario1))
		print2('Unsuccessfully saved scenario error code: {}'.format(sce))
	else:
		print1('No Scenario Active to Save.........................................', bf=2, af=0)
	return None


def get_active_variations():			# Get Active Network Variations
	"""
		Get active variations
	:return list variations: Returns list of variations currently active
	"""
	variations =  app.GetActiveNetworkVariations()
	print1('Current Active Variations: ', bf=2, af=0)
	if len(variations) > 1:
		for item in variations:
			aa = str(item)
			pp = aa.split("Variations.IntPrjfolder\\")
			ss = pp[1]
			tt = ss.split(".IntScheme")
			print1(tt[0], bf=1, af=0)
	elif len(variations) == 1:
		print1(variations, bf=1, af=0)
	else:
		print1('No Variations Active', bf=1, af=0)
	return variations


def create_variation(folder, pfclass, name):
	"""
		Create a new variaiton
	:param str folder: Name of power factory folder variation should be saved in
	:param pfclass: Class of variation to be created
	:param str name: Name for variation
	:return powerfactory.Variation: Handle for newly created variation
	"""
	variation = create_object(folder, pfclass, name)
	# Change color of variation
	variation.icolor = 1
	print1('Variation {} created'.format(variation))
	# #app.PrintPlain(variation)
	return variation


def activate_variation(variation): 		# Activate Scenario
	""" 
		Activate previously created variation
	:param powerfactory.Variation variation: handle to existing powerfactory variation
	:return int sce: Integer (0,1) on whether success or fail on activating variation
	"""
	sce = variation.Activate() 											# Activate Study case
	if sce == 0:
		print1('Activated Variation Successfully: {}'.format(variation), bf=1, af=0)
	elif sce == 1:
		print2('Error Unsuccessfully Activated Variation: {}........................'.format(variation))
		print2('Unsuccessfully Activated Variation Error Code: {}'.format(sce))
	return sce


def create_stage(location, pfclass, name):
	"""
		Creates a new stage in powerfactory
	:param powerfactory.Location location: Handle to powerfacory location
	:param str pfclass: String describing the powerfactory stage to be created
	:param ztr name: Name of new stage to be created
	:return powerfactory.Stage stage: Handle to newly created powerfactory stage
	"""
	stage = location.CreateObject(pfclass, name)
	stage.loc_name = name
	activate_stage(stage)
	return stage


def activate_stage(stage):
	"""
		Activate stage created by PowerFactory
	:param powerfactory.Stage stage: Handle to powerfactory Stage to be activated
	:return: None
	"""
	sce = stage.Activate()
	if sce == 0:
		print1('Activated Variation Stage Successfully: {}'.format(stage), bf=1, af=0)
	elif sce != 0:
		print2('Error Unsuccessfully Activated Variation Stage: {}........................'.format(stage))
		print2('Unsuccessfully Activated Variation Stage Error Code: {}'.format(sce))
	return None


def load_flow(load_flow_settings):		# Inputs load flow settings and executes load flow
	"""
		Run load flow in powerfactory
	:param list load_flow_settings: List of settings for powerfactory when running loadflow 
	:return int error_code: Error code provided by powerfactory determining its success 
	"""
	# TODO: Setting should only need setting once rather than every time load_flow is run so could be defined in
	# TODO: + constants
	t1 = time.clock()
	## Loadflow settings
	## -------------------------------------------------------------------------------------
	# Basic
	ldf.iopt_net = load_flow_settings[0]          		# Calculation method (0 Balanced AC, 1 Unbalanced AC, DC)
	ldf.iopt_at = load_flow_settings[1]            		# Automatic Tap Adjustment
	ldf.iopt_ashnt = load_flow_settings[2]        		# Automatic Shunt Adjustment
	ldf.iopt_lim = load_flow_settings[3]             	# Consider Reactive Power Limits
	ldf.iopt_ashnt = load_flow_settings[4]             	# Consider Reactive Power Limits Scaling Factor
	ldf.iopt_tem = load_flow_settings[5]               	# Temperature Dependency: Line Cable Resistances (0 ...at 20C, 1 at Maximum Operational Temperature)
	ldf.iopt_pq = load_flow_settings[6]               	# Consider Voltage Dependency of Loads
	ldf.iopt_fls = load_flow_settings[7]               	# Feeder Load Scaling
	ldf.iopt_sim = load_flow_settings[8]              	# Consider Coincidence of Low-Voltage Loads
	ldf.scPnight = load_flow_settings[9]            	# Scaling Factor for Night Storage Heaters

	# Active Power Control
	ldf.iopt_apdist = load_flow_settings[10]           	# Active Power Control (0 as Dispatched, 1 According to Secondary Control,
															# 2 Acording to Primary Control, 3 Acording to Inertias)
	ldf.iopt_plim = load_flow_settings[11]            	# Consider Active Power Limits
	ldf.iPbalancing = load_flow_settings[12]          	# (0 Ref Machine, 1 Load, Static Gen, Dist slack by loads, Dist slack by Sync,
	# ldf.rembar = load_flow_settings[13] # Reference Busbar
	ldf.phiini = load_flow_settings[14]         		# Angle

	# Advanced Options
	ldf.i_power = load_flow_settings[15]               	# Load Flow Method ( NR Current, 1 NR (Power Eqn Classic)
	ldf.iopt_notopo = load_flow_settings[16]          	# No Topology Rebuild
	ldf.iopt_noinit = load_flow_settings[17]          	# No initialisation
	ldf.utr_init = load_flow_settings[18]           	# Consideration of transformer winding ratio
	ldf.maxPhaseShift = load_flow_settings[19]      	# Max Transformer Phase Shift
	ldf.itapopt = load_flow_settings[20]               	# Tap Adjustment ( 0 Direct, 1 Step)
	ldf.krelax = load_flow_settings[21]              	# Min COntroller Relaxation Factor

	ldf.iopt_stamode = load_flow_settings[22]        	# Station Controller (0 Standard, 1 Gen HV, 2 Gen LV
	ldf.iopt_igntow = load_flow_settings[23]          	# Modelling Method of Towers (0 With In/ Output signals, 1 ignore couplings, 2 equation in lines)
	ldf.initOPF = load_flow_settings[24]            	# Use this load flow for initialisation of OPF
	ldf.zoneScale = load_flow_settings[25]            	# Zone Scaling ( 0 Consider all loads, 1 Consider adjustable loads only)

	# Iteration Control
	ldf.itrlx = load_flow_settings[26]                	# Max No Iterations for Newton-Raphson Iteration
	ldf.ictrlx = load_flow_settings[27]               	# Max No Iterations for Outer Loop
	ldf.nsteps = load_flow_settings[28]               	# Max No Iterations for Number of steps

	ldf.errlf = load_flow_settings[29]             	   	# Max Acceptable Load Flow Error for Nodes
	ldf.erreq = load_flow_settings[30]             		# Max Acceptable Load Flow Error for Model Equations
	ldf.iStepAdapt = load_flow_settings[31]       		# Iteration Step Size ( 0 Automatic, 1 Fixed Relaxation)
	ldf.relax = load_flow_settings[32]             		# If Fixed Relaxation factor
	ldf.iopt_lev = load_flow_settings[33]         		# Automatic Model Adaptation for Convergence 

	# Outputs
	ldf.iShowOutLoopMsg = load_flow_settings[34] 		# Show 'outer Loop' Messages
	ldf.iopt_show = load_flow_settings[35]       		# Show Convergence Progress Report
	ldf.num_conv = load_flow_settings[36]         		# Number of reported buses/models per iteration
	ldf.iopt_check = load_flow_settings[37]      		# Show verification report
	ldf.loadmax = load_flow_settings[38]           		# Max Loading of Edge Element
	ldf.vlmin = load_flow_settings[39]            		# Lower Limit of Allowed Voltage
	ldf.vlmax = load_flow_settings[40]             		# Upper Limit of Allowed Voltage
	# ldf.outcmd =  load_flow_settings[41]          		# Output
	ldf.iopt_chctr = load_flow_settings[42]    			# Check Control Conditions
	# ldf.chkcmd = load_flow_settings[43]            	# Command

	# Load Generation Scaling
	ldf.scLoadFac = load_flow_settings[44]          	# Load Scaling Factor
	ldf.scGenFac = load_flow_settings[45]              	# Generation Scaling Factor
	ldf.scMotFac = load_flow_settings[46]              	# Motor Scaling Factor

	# Low Voltage Analysis
	ldf.Sfix = load_flow_settings[47]                  	# Fixed Load kVA
	ldf.cosfix = load_flow_settings[48]                	# Power Factor of Fixed Load
	ldf.Svar = load_flow_settings[49]                  	# Max Power Per Customer kVA
	ldf.cosvar = load_flow_settings[50]                	# Power Factor of Variable Part
	ldf.ginf = load_flow_settings[51]                  	# Coincidence Factor
	ldf.i_volt = load_flow_settings[52]          		# Voltage Drop Analysis (0 Stochastic Evaluation, 1 Maximum Current Estimation)

	# Advanced Simulation Options
	ldf.iopt_prot = load_flow_settings[53]        		# Consider Protection Devices ( 0 None, 1 all, 2 Main, 3 Backup)
	ldf.ign_comp = load_flow_settings[54]             	# Ignore Composite Elements

	print2('DEBUG - Load flow started')
	error_code = ldf.Execute()
	print2('DEBUG - Load flow completed')
	t2 = time.clock() - t1
	if error_code == 0:
		print1('Load Flow calculation successful, time taken: {:.2f} seconds'.format(t2), bf=1, af=0)
	elif error_code == 1:
		print2('Load Flow failed due to divergence of inner loops, time taken: {:.2f} seconds..............'.format(t2))
	elif error_code == 2:
		print2('Load Flow failed due to divergence of outer loops, time taken: {:.2f} seconds..............'.format(t2))
	return error_code


def harm_load_flow(results, harmonic_loadflow_settings):		# Inputs load flow settings and executes load flow
	"""
		Runs harmonic load flow
	:param results: Results variable provided as an input to the powerfactory harmonic load flow
	:param list harmonic_loadflow_settings: Harmonic load flow settings
	:return int error_code: Error code returned by harmonic load flow 
	"""
	t1 = time.clock()
	## Loadflow settings
	## -------------------------------------------------------------------------------------
	# Basic
	hldf.iopt_net = harmonic_loadflow_settings[0]               	# Calculation method (0 Balanced AC, 1 Unbalanced AC, DC)
	hldf.iopt_allfrq = harmonic_loadflow_settings[1]				# Calculate Harmonic Load Flow 0 - Single Frequency 1 - All Frequencies
	hldf.iopt_flicker = harmonic_loadflow_settings[2] 				# Calculate Flicker
	hldf.iopt_SkV = harmonic_loadflow_settings[3] 					# Calculate Sk at Fundamental Frequency
	hldf.frnom = harmonic_loadflow_settings[4]            			# Nominal Frequency
	hldf.fshow = harmonic_loadflow_settings[5]             			# Output Frequency
	hldf.ifshow = harmonic_loadflow_settings[6]  					# Harmonic Order
	hldf.p_resvar = results          								# Results Variable
	# hldf.cbutldf =  harmonic_loadflow_settings[8]               	# Load flow
	
	# IEC 61000-3-6
	hldf.iopt_harmsrc = harmonic_loadflow_settings[9]				# Treatment of Harmonic Sources
	
	# Advanced Options
	hldf.iopt_thd = harmonic_loadflow_settings[10] 					# Calculate HD and THD 0 Based on Fundamental Frequency values 1 Based on rated voltage/current
	hldf.maxHrmOrder = harmonic_loadflow_settings[11] 				# Max Harmonic order for calculation of THD and THF
	hldf.iopt_HF = harmonic_loadflow_settings[12] 					# Calculate Harmonic Factor (HF)
	hldf.ioutall = harmonic_loadflow_settings[13] 					# Calculate R, X at output frequency for all nodes
	hldf.expQ = harmonic_loadflow_settings[14] 						# Calculation of Factor-K (BS 7821) for Transformers
	
	error_code = hldf.Execute()
	t2 = time.clock() - t1
	if error_code == 0:
		print1('Harmonic Load Flow calculation successful: {:.2f} seconds'.format(t2), bf=1, af=0)
	elif error_code > 0:
		print2('Harmonic Load Flow calculation unsuccessful: {:.2f} seconds...............................'.format(t2))
	return error_code


def freq_sweep(results, fsweep_settings):		# Inputs Frequency Sweep Settings and executes sweep
	"""
		Sets up and runs frequency sweep
	:param results: powerfactory results variable
	:param list fsweep_settings: Settings for frequency sweep
	:return int error_code: Error code showing whether frequency sweep was successful 
	"""
	t1 = time.clock()
	## Frequency Sweep Settings
	## -------------------------------------------------------------------------------------
	# Basic
	# nomfreq and maxfrq reported as not being used
	# TODO: Check whether input frq.frnom should actually be nomfreq
	# nomfreq = fsweep_settings[0]                  # Nominal Frequency
	# maxfrq = fsweep_settings[1]                 	# Maximum Frequency
	frq.iopt_net = fsweep_settings[2]                # Network Representation (0=Balanced 1=Unbalanced)
	frq.fstart = fsweep_settings[3]                	# Impedance Calculation Start frequency
	frq.fstop = fsweep_settings[4]              # Stop Frequency
	frq.fstep = fsweep_settings[5]                 # Step Size
	frq.i_adapt = fsweep_settings[6]                 # Automatic Step Size Adaption
	frq.frnom = fsweep_settings[7]             # Nominal Frequency
	frq.fshow = fsweep_settings[8]              # Output Frequency
	frq.ifshow = fsweep_settings[9]   # Harmonic Order
	frq.p_resvar = results          # Results Variable
	# frq.cbutldf = fsweep_settings[11]                 # Load flow

	# Advanced
	frq.errmax = fsweep_settings[12]               # Setting for Step Size Adaption    Maximum Prediction Error
	frq.errinc = fsweep_settings[13]              # Minimum Prediction Error
	frq.ninc = fsweep_settings[14]                   # Step Size Increase Delay
	frq.ioutall = fsweep_settings[15]                 # Calculate R, X at output frequency for all nodes

	error_code = frq.Execute()	
	t2 = time.clock() - t1
	if error_code == 0:
		print1('Frequency Sweep calculation successful, time taken: {:.2f} seconds'.format(t2), bf=1, af=0)
	elif error_code > 0:
		print2('Frequency Sweep calculation unsuccessful, time taken: {:.2f} seconds.......................'.format(t2))
	return error_code


def switch_coup(element, service):			# Switches an Coupler out if 0 in if 1
	"""
		Changes status of coupler (i.e. 0 if in or 1 if out)
	:param powerfacotry.Element element: Handle to powerfactory element to have status changed 
	:param service: 
	:return: None
	"""
	element.on_off = service
	if service == 0:
		print1('Switching Element: {} Out of service "'.format(element), bf=1, af=0)
	if service == 1:
		print1('Switching Element: {} In to service '.format(element), bf=1, af=0)
	return None


def check_if_folder_exists(location, name):		# Checks if the folder exists
	"""
		Check if power factory folder already exists
	:param powerfacotry.Location location: Handle to existing powerfactory location 
	:param str name: Name of folder 
	:return (powerfactory.FolderObject, status), (new_object, folder_exists): Handle to folder and status on whether it already exists 
	"""
	_new_object = location.GetContents('{}.IntFolder'.format(name))
	folder_exists = 0
	if len(_new_object) > 0:
		print1('Folder already exists: {}'.format(name), bf=2, af=0)
		folder_exists = 1
	return _new_object, folder_exists


def create_folder(location, name):		# Creates Folder in location
	"""
		Create folder in new location
	:param powerfactory.Location location: PowerFactory location that folder should be created in 
	:param str name: Name of folder to be created 
	:return _new_object, status: Handle for new_object and True/False status on whether it already exists 
	"""
	# _new_object used instead of new_object to avoid showing
	print1('Creating Folder: {}'.format(name), bf=1, af=0)
	folder1, folder_exists = check_if_folder_exists(location, name)				# Check if the Results folder exists if it doesn't create it using date and time
	if folder_exists == 0:
		_new_object = location.CreateObject("IntFolder",name)
		# loc_name = name							# Name of Folder
		# owner = "Barry"							# Owner
		# iopt_sys = 0							# Attributes System
		# iopt_type = 0							# Folder Type 0 Common
		# for_name = ""							# Foreign key
		# desc = ""								# Description
	else:
		_new_object = folder1[0]
	return _new_object, folder_exists


# Creates a mutual Impedance list from the terminal list in a folder under the active studycase
def create_mutual_impedance_list(location, terminal_list):
	"""
		Create a mutual impedance list from the terminal list in a folder under the active studycase
	:param powerfacory.Location location: Handle for location to be created 
	:param list terminal_list: List of terminals 
	:return list list_of_mutual: List of mutual impedances 
	"""
	print1('Creating: Mutual Impedance List of Terminals', bf=1, af=0)
	terminal_list1 = list(terminal_list)
	list_of_mutual = []
	# # count = 0
	# TODO: Improve since this is a loop of loops
	for y in terminal_list1:
		for x in terminal_list1:
			if x[3] != y[3]:
				name = '{}_{}'.format(y[0],x[0])
				elmmut = create_mutual_elm(location, name, y[3], x[3])
				list_of_mutual.append([str(y[0]), name, elmmut, y[3], x[3]])
	return list_of_mutual


def create_mutual_elm(location, name, bus1, bus2):		# Creates Mutual Impedance between two terminals
	"""
		Create mutual impedance between two terminals
	:param powerfactory.Location location: Handle for location to save mutual impedance 
	:param str name: Name for mutual impedance 
	:param bus1: Terminal 1 of mutual impedance
	:param bus2: Terminal 2 of mutual impedance
	:return: PowerFactory.ElmMut elmmut: Handle for mutual impedance
	"""
	# elmmut = app.GetFromStudyCase(name + )				# Get relevant object or create if it doesn't exist
	elmmut = create_object(location, "ElmMut", name)
	elmmut.loc_name = name
	elmmut.bus1 = bus1
	elmmut.bus2 = bus2
	elmmut.outserv = 0
	return elmmut


def get_object(object_to_retrieve):			# retrieves an object based on filter strings
	"""
		Retrieves an object based on filter strings
	:param str object_to_retrieve: Name of object to be returned 
	:return powerfactory.Object obj: Handle for object returns 
	"""
	ob1 = app.GetCalcRelevantObjects(object_to_retrieve)
	return ob1


def delete_object(object_to_delete):			# retrieves an object based on filter strings
	ob1 = object_to_delete.Delete()
	if ob1 == 0:
		print1('Object Successfully Deleted: {}'.format(object_to_delete), bf=1, af=0)
	else:
		print2('Error deleting object: {}.....................'.format(object_to_delete))
	return None


def check_if_object_exists(location, name):  	# Check if the object exists
	"""
		Check if an object exists by name
	:param powerfactory.Location location: Handle for PF location to investigate 
	:param str name: Name of object to look for 
	:return object_exists, new_object: True / False on whether object exists, handle for powerfactory.Object 
	"""
	print1('{} {}'.format(location, name), bf=2, af=0)
	#_new_object used instead of new_object to avoid shadowing
	_new_object = location.GetContents(name)
	object_exists = 0
	if len(_new_object) > 0:
		print1('Object Exists: {}'.format(name), bf=2, af=0)
		object_exists = 1
	return object_exists, _new_object


def add_copy(folder, object, name1):		# copies an object to a new folder Name 1 = new name
	"""
		Copies an object to a new folder
	:param powerfactory.Folder folder: Folder in which object should be copied 
	:param powerfactor.Object object: Object to be copied 
	:param str name1: Name of new object 
	:return: 
	"""
	new_object = folder.AddCopy(object, name1)
	if new_object is not None:
		print1('Copying object {} successful'.format(object), bf=1, af=0)
	else:
		print2('Error AddCopy Unsuccessful: {} to {} as {}'.format(object, folder, name1))
	return new_object


def create_object(location, pfclass, name):			# Creates a database object in a specified location of a specified class
	"""
		Creates a database object in a specified location of a specified class
	:param powerfactory.Location location: Location in which new object should be created 
	:param str pfclass: Type of element to be created 
	:param str name: Name to be given to new object 
	:return powerfactory.Object _new_object: Handle to object returns 
	"""
	# _new_object used instead of new_object to avoid shadowing
	_new_object = location.CreateObject(pfclass, name)
	return _new_object


def create_results_file(location, name, type_of_file):			# Creates Results File
	"""
		Creates a suitale results file to store the frequency / harmonic results
	:param powerfactory.Location location: Handle for location into which to store the results 
	:param str name: Name of results file 
	:param type_of_file: Type of file (Frequency / Harmonic) 
	:return powerfactory.results sweep: Handle for results file 
	"""
	# Manipulating Results Files
	sweep = create_object(location, "ElmRes", name)
	_ = sweep.Clear()								# Clears Data
	variable_contents = sweep.GetContents()			# Gets the existing variables
	for cc in variable_contents:					# Loops through and deletes the existing variables
		cc.Delete()
	sweep.calTp = type_of_file						# Frequency / Harmonic
	# TODO: See if sweep.header and sweep.desc are still required
	sweep.header = ["Hello Barry"]
	sweep.desc = ["Barry Description"]
	return sweep


def check_list_of_studycases(list_to_check):		# Check List of Projects, Study Cases, Operational Scenarios,
	"""
		Check list of projects, study cases, operational scenarios, etc. solve for load flow
		Produces a new operational scenario and study case of each contingency so that each study case can be split
		out into separate parallel processing functions
	:param list list_to_check: List of items to check 
	:return dict prj_dict:  Dictionary of projects where the study cases associated with each project are contained within 
	"""
	# TODO: Check function since there are a lot of unresolved references
	print1('___________________________________________________________________________________________________', bf=2, af=0)
	print1(('Checking all Projects, Study Cases and Scenarios Solve for Load Flow, it will also check N-1 and create ' +
		 'the operational scenarios if Pre_Case_Check is True\n'),
		   bf=2, af=0)
	# _count_studycase used instead of count_studycase to avoid shadowing
	_count_studycase = 0
	new_list =[]

	# Empty list which will be populated with the new classes
	cls_list = []
	prj_dict = dict()
	while _count_studycase < len(list_to_check):
		# ERROR: Previously was not actually looking at the list passed to function
		# # prj = activate_project(List_of_Studycases[_count_studycase][1])												# Activate Project
		# TODO: Efficieny - This is activating a new project even if it is the same
		project_name = list_to_check[_count_studycase][1]
		_prj = activate_project(project_name)  # Activate Project

		if len(str(_prj)) > 0:
			study_case, _study_error = activate_study_case(list_to_check[_count_studycase][2])									# Activate Case
			if _study_error == 0:
				scenario, scen_err = activate_scenario(list_to_check[_count_studycase][3])										# Activate Scenario
				if scen_err == 0:
					print2('Load flow being run for study case {}'.format(list_to_check[_count_studycase]))
					ldf_err = load_flow(Load_Flow_Setting)																			# Perform Load Flow
					print2('DEBUG - Load flow study completed with error code {}'.format(ldf_err))					
					if ldf_err == 0 or Skip_Unsolved_Ldf == False:
						new_list.append(list_to_check[_count_studycase])

						print1("Studycase Scenario Solving added to analysis list: " + str(list_to_check[_count_studycase]),
							   bf=2, af=0)

						# TODO: At this point could create just a list that references the newly created study case and return that,
						# TODO: The newly created study cases can then just be activated and deactivated as appropriate.

						if Pre_Case_Check:																	# Checks all the contingencies and terminals are in the prj,cas
							new_contingency_list, con_ok = check_contingencies(List_of_Contingencies) 				# Checks to see all the elements in the contingency list are in the case file
							terminals_index, term_ok = check_terminals(List_of_Points)								# Checks to see if all the terminals are in the case file skips any that aren't
							# Adjusted to create new study_case for each op_scenario

							study_case_folder = app.GetProjectFolder('study')
							study_case_results_folder, _folder_exists2 = create_folder(study_case_folder,
																					   Operation_Scenario_Folder)

							operation_case_folder = app.GetProjectFolder("scen")
							_op_sc_results_folder, _folder_exists2 = create_folder(operation_case_folder,
																				   Operation_Scenario_Folder)

							# Create a results folder for each project so that the reference can be included in the study_cls
							results_folder, folder_exists = create_folder(_prj, Results_Folder)

							# Create ComTasks and store in parent_study_case_folder (required location)
							task_automation = create_object(study_case_results_folder, 'ComTasks',
															'Task_Automation_{}'.format(start1))

							cont_count = 0
							while cont_count < len(new_contingency_list):
								print1('Carrying out Contingency Pre Stage Check: {}'.format(new_contingency_list[cont_count][0]),
									   bf=2, af=0)
								deactivate_scenario()																# Can't copy activated Scenario so deactivate it
								# Can't copy activated studu case so deactivate it
								deactivate_study_case()
								cont_name = '{}_{}'.format(List_of_Studycases[_count_studycase][0],
														   new_contingency_list[cont_count][0])
								_new_study_case = add_copy(study_case_results_folder,
														   study_case,
														   cont_name)
								# #print2(_new_study_case)
								_new_scenario = add_copy(_op_sc_results_folder, scenario,
														 cont_name)	# Copies the base scenario
								_new_study_case.Activate()
								_ = activate_scenario1(_new_scenario)										# Activates the base scenario
								if new_contingency_list[cont_count][0] != "Base_Case":								# Apply Contingencies if it is not the base case
									# Take outages described for contingency
									for _switch in new_contingency_list[cont_count][1:]:
										switch_coup(_switch[0], _switch[1])

								save_active_scenario()

								# TODO: This load flow may be unnecessary
								_ = load_flow(Load_Flow_Setting)															# Carry out load flow

								# Deactivate new study case and reactivate old study case
								_new_study_case.Deactivate()
								study_case.Activate()

								# Create new class reference with all the details for this contingency and then add to
								# list to be returned
								_study_cls = hast.pf.PFStudyCase(full_name=cont_name,
																 list_parameters=list_to_check[_count_studycase],
																 cont_name=new_contingency_list[cont_count][0],
																 sc=_new_study_case,
																 op=_new_scenario,
																 prj=_prj,
																 res_folder=results_folder,
																 task_auto=task_automation,
																 uid=start1)
								# Add study case to dictionary of projects
								if project_name not in prj_dict.keys():
									prj_dict[project_name] = hast.pf.PFProject(name=project_name,
																			   prj=_prj,
																			   res_folder=results_folder,
																			   task_auto=task_automation)

								# Add study case to file
								prj_dict[project_name].sc_cases.append(_study_cls)


								# #cls_list.append(_study_cls)
								# TODO: Use enumerator rather than iterating counter
								cont_count = cont_count + 1
					else:
						print2("Problem with Loadflow: " + str(list_to_check[_count_studycase][0]))
				else:
					print2("Problem with Scenario: " + str(list_to_check[_count_studycase][0]) + " " + str(list_to_check[_count_studycase][3]))
			else:
				print2('Problem with Studycase: {} {}'
					   .format(list_to_check[_count_studycase][0], list_to_check[_count_studycase][2]))
		else:
			print2('Problem Activating Project: {} {}'
				   .format(list_to_check[_count_studycase][0], list_to_check[_count_studycase][1]))
		_count_studycase += 1
	print1("Finished Checking Study Cases", bf=1, af=0)
	print1("___________________________________________________________________________________________________",
		   bf=2,af=2)
	# #return new_list
	# Get unique list of projects
	return prj_dict


def check_terminals(list_of_points): 		# This checks and creates the list of terminals to add to the Results file
	"""
		Creates list of terminals to be added to the results file
	:param list list_of_points: 
	:return: (list, bool) (terminals_index, terminals_ok): list of terminal indexes, success on adding terminals
	"""
	terminals_ok = 0
	terminals_index = []														# Where the calculated variables will be stored
	tm_count = 0
	while tm_count < len(list_of_points):										# This loops through the variables adding them to the results file
		t = app.GetCalcRelevantObjects(list_of_points[tm_count][1])				# Finds the Substation
		if len(t) == 0:															# If it doesn't find it it reports it and skips it
			print2("Python substation entry for does not exist in case: " + list_of_points[tm_count][1] + "..............................................")
		else:
			t1 = t[0].GetContents()													# Gets the Contents of the substations (ie objects) 
			terminal_exists = False
			for t2 in t1:															# Gets the contents of the objects in the Substaion
				if list_of_points[tm_count][2]  in str(t2):												# Checks to see if the terminal is there
					terminals_index.append([list_of_points[tm_count][0],
											list_of_points[tm_count][1],
											list_of_points[tm_count][2],
											t2])					# Appends Terminals ( Name, Terminal Name, Terminal object data)
					terminal_exists = True											# Marks that it found the terminal
			if not terminal_exists:
				print2("Python Entry does not exist in case: " + list_of_points[tm_count][2] + ".ElmTerm ..............................................")
				terminals_ok = 1
		tm_count = tm_count + 1
	print1("Terminals Used for Analysis: ", bf=2, af=0)
	tm_count = 0
	while tm_count < len(list_of_points):
		print1(list_of_points[tm_count], bf=1, af=0)
		tm_count = tm_count + 1
	return terminals_index, terminals_ok


def check_contingencies(list_of_contingencies): 		# This checks and creates the list of terminals to add to the Results file
	"""
		Checks the status of each contingency
	:param list list_of_contingencies: List of contingencies to be checked 
	:return: List of ocntinencies to actually be studied 
	"""
	contingencies_ok = 0
	new_contingency_list = []															# Where the calculated variables will be stored
	for item in list_of_contingencies:													# This loops through the contingencies to find the couplers
		skip_contingency = False
		list_of_couplers = []
		if item[0] == "Base_Case":														# Skips the base case
			# # NOT USED - coupler_exists = True
			list_of_couplers.append("Base_Case")
			list_of_couplers.append(0)
		else:
			list_of_couplers.append(item[0])
			for aa in item[1:]:
				coupler_exists = False
				t = app.GetCalcRelevantObjects(aa[0] + ".ElmSubstat")					# Finds the Substation 
				if len(t) == 0:															# If it doesn't find it it stops the script
					print2("Contingency entry: " + item[0] + ". Substation does not exist in case: " + aa[0] + ".ElmSubstat Check Python Entry..............................................")
					print2("Skipping Contingency")
					skip_contingency = True
				else:
					t1 = t[0].GetContents()													# Gets the Contents of the substations (ie objects) 
					for t2 in t1:															# Gets the contents of the objects in the Substaion
						if ".ElmCoup" in str(t2):											# Filters for Terminals
							if aa[1] in str(t2):											# Filters for items in terminals
								if aa[2] == "Open":
									breaker_operation = 0
									list_of_couplers.append([t2,
															 breaker_operation])  # Appends Terminals ( Name, Terminal Name, Terminal object data)
								elif aa[2] == "Close":
									breaker_operation = 1
									list_of_couplers.append([t2,
															 breaker_operation])  # Appends Terminals ( Name, Terminal Name, Terminal object data)
								else:
									print2("Contingency entry: " + item[0] + ". Coupler in Substation: " + aa[0] +  " " + aa[1] + " could not carry out: " + aa[2] + " ..............................................")

						coupler_exists = True										# Marks that it found the terminal		
					if not coupler_exists:
						print2("Contingency entry: " + item[0] + ". Coupler does not exist in Substation: " + aa[0] +  " " + aa[1] + " ..............................................")
						print2("Skipping Contingency")
						skip_contingency = True
		if skip_contingency:
			contingencies_ok = 1
		elif not skip_contingency:
			new_contingency_list.append(list_of_couplers)
	print1("Contingencies Used for Analysis:", bf=2, af=0)
	for item in new_contingency_list:
		print1(item, bf=1, af=0)
	return new_contingency_list, contingencies_ok


def add_vars_res(elmres, element, res_vars):	# Adds the results variables to the results file
	"""
		Adds the results variables to the results file
	:param elmres: 
	:param element: 
	:param res_vars: 
	:return: None
	"""
	if len(res_vars) > 1:
		for x in res_vars:
			elmres.AddVariable(element, x)
	elif len(res_vars) == 1:
		elmres.AddVariable(element,res_vars[0])
	return None


def retrieve_results(elmres, res_type):			# Reads results into python lists from results file
	"""
		Reads results into python lists from results file for processing to add to Excel
	:param powerfactory.Results elmres: handle for powerfactory results file 
	:param int res_type: Type of results being dealt with 
	:return: 
	"""
	# Note both column number and row start at 1.
	# The first column is usually the scale ie timestep, frequency etc.
	# The columns are made up of Objects from left to right (ElmTerm, ElmLne)
	# The Objects then have sub variables (m:R, m:X etc)
	elmres.Load()
	cno = elmres.GetNumberOfColumns()	# Returns number of Columns 
	rno = elmres.GetNumberOfRows()		# Returns number of Rows in File
	results = []
	for i in range(cno):
		column = []
		p = elmres.GetObject(i) 		# Object
		d = elmres.GetVariable(i)		# Variable
		column.append(d)
		column.append(str(p))
		# column.append(d)
		# app.PrintPlain([i,p,d])	
		for j in range(rno):
			r, t = elmres.GetValue(j, i)
			# app.PrintPlain([i,p,d,j,t])
			column.append(t)
		results.append(column)
	if res_type == 1:
		results = results[:-1]
	scale = results[-1:]
	results = results[:-1]
	elmres.Release()
	return scale[0], results


# Main Engine ------------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------------
# Following if statement stops the code being run unless it is the main script
if __name__ == '__main__':
	""" Ensures this code is only run when run as the main script and not for unittesting """
	# TODO: If want to unittest PF will need to put this into a function
	DIG_PATH = """C:\\Program Files\\DIgSILENT\\PowerFactory 2016 SP3\\"""
	sys.path.append(DIG_PATH)
	os.environ['PATH'] = os.environ['PATH'] + ';' + DIG_PATH
	Title = ("""::::::::::::::::::::::::::::::::::::::::::::::::::::::::::\n
		NAME:             HAST Harmonics Automated Simulation Tool\n
		VERSION:          Development Verson by David Mills (PSC)\n
		AUTHOR:           Barry O'Connell\n
		::::::::::::::::::::::::::::::::::::::::::::::::::::::::::\n""")

	# File location of this script when running
	filelocation = os.getcwd() + "\\"

	# Start Timer
	start = time.clock()
	start1 = (time.strftime("%y_%m_%d_%H_%M_%S"))

	# Power factory commands
	# --------------------------------------------------------------------------------------------------------------------
	app = powerfactory.GetApplication()  # Start PowerFactory  in engine  mode
	# help("powerfactory")
	user = app.GetCurrentUser()  # Get the current active user
	ldf = app.GetFromStudyCase("ComLdf")  # Get load flow command
	hldf = app.GetFromStudyCase("ComHldf")  # Get Harmonic load flow
	frq = app.GetFromStudyCase("ComFsweep")  # Get Frequency Sweep Command
	ini = app.GetFromStudyCase("ComInc")  # Get Dynamic Initialisation
	sim = app.GetFromStudyCase("ComSim")  # Get Dynamic Simulation
	shc = app.GetFromStudyCase("ComShc")  # Get short circuit command
	res = app.GetFromStudyCase("ComRes")  # Get Result Export Command
	wr = app.GetFromStudyCase("ComWr")  # Get Write command for wmf and bmp files
	app.ClearOutputWindow()  # Clear Output Window

	Error_Count = 1

	# Enter what Variables you want to look at for terminals
	FS_Terminal_Variables = ["m:R", "m:X", "m:Z", "m:phiz"]
	Mutual_Variables = ["c:Z_12"]
	HRM_Terminal_Variables = ["m:HD"]
	# Import Excel
	Import_Workbook = filelocation + "HAST_V1_2_Inputs.xlsx"					# Gets the CWD current working directory
	Variation_Name = "Temporary_Variation" + start1

	# Create excel instance to deal with retrieving import data from excel
	with hast.excel_writing.Excel(print_info=print1, print_error=print2) as excel_cls:
		analysis_dict = excel_cls.import_excel_harmonic_inputs(workbookname=Import_Workbook) 			# Reads in the Settings from the spreadsheet

	Study_Settings = analysis_dict["Study_Settings"]
	if len(Study_Settings) != 20:
		print2('Error, Check input Study Settings there should be 20 Items in the list there are only: {} {}'
			   .format(len(Study_Settings), Study_Settings))
	if not Study_Settings[0]:											# If there is no output location in the spreadsheet it sets it to the CWD
		Results_Export_Folder = filelocation
	else:
		Results_Export_Folder = Study_Settings[0]								# Folder to Export Excel Results too

	# Declare file names
	Excel_Results = Results_Export_Folder + Study_Settings[1] + start1			# Name of Exported Results File
	Progress_Log = Results_Export_Folder + Study_Settings[2] + start1 + ".txt"	# Progress File
	Error_Log = Results_Export_Folder + Study_Settings[3] + start1 + ".txt"		# Error File
	Debug_Log = Results_Export_Folder + 'DEBUG' + start1 + '.txt'

	# Setup logger with reference to powerfactory app
	logger = hast.logger.Logger(pth_debug_log=Debug_Log,
								pth_progress_log=Progress_Log,
								pth_error_log=Error_Log,
								app=app)

	# Random_Log = Results_Export_Folder + "Random_Log_" + start1 + ".txt"		# For printing random info solely for development
	Net_Elm = Study_Settings[4]													# Where all the Network elements are stored
	Mut_Elm_Fld = Study_Settings[5] + start1									# Name of the folder to create under the network elements to store mutual impedances
	Results_Folder = Study_Settings[6] + start1									# Name of the folder to keep results under studycase
	Operation_Scenario_Folder = Study_Settings[7]	+ start1 					# Name of the folder to store Operational Scenarios
	Pre_Case_Check = Study_Settings[8]											# Checks the N-1 for load flow convergence and saves operational scenarios.
	FS_Sim = Study_Settings[9]													# Carries out Frequency Sweep Analysis
	HRM_Sim = Study_Settings[10]												# Carries out Harmonic Load Flow Analysis
	Skip_Unsolved_Ldf = Study_Settings[11]										# Skips the frequency sweep if the load flow doesn't solve
	Delete_Created_Folders = Study_Settings[12]									# Deletes the Results folder, Mutual Elements and the Operational Scenario folder
	Export_to_Excel = Study_Settings[13]										# Export the results to Excel
	Excel_Visible = Study_Settings[14]											# Makes Excel Visible while plotting, Can be annoying if you are doing other work as if you click the excel screen it stops the simulation
	Excel_Export_RX = Study_Settings[15]										# Export RX and graph the Impedance Loci in Excel
	Excel_Convex_Hull = Study_Settings[16]										# This calculates the minimum points for the Loci
	Excel_Export_Z = Study_Settings[17]											# Graph the Frequency Sweeps in Excel
	Excel_Export_Z12 = Study_Settings[18]										# Export Mutual Impedances to excel
	Excel_Export_HRM = Study_Settings[19]										# Export Harmonic Data
	print1(Title, bf=1, af=0)
	for keys,values in analysis_dict.items():									# Prints all the inputs to progress log
		print1(keys, bf=1, af=0)
		print1(values, bf=1, af=0)
	List_of_Studycases = analysis_dict["Base_Scenarios"]						# Uses the list of Studycases
	if len(List_of_Studycases) <1:												# Check there are the right number of inputs
		print2("Error - Check excel input Base_Scenarios there should be at least 1 Item in the list ")
	List_of_Contingencies = analysis_dict["Contingencies"]						# Uses the list of Contingencies
	if len(List_of_Contingencies) <1:											# Check there are the right number of inputs
		print2("Error - Check excel input Contingencies there should be at least 1 Item in the list ")
	List_of_Points = analysis_dict["Terminals"]									# Uses the list of Terminals
	if len(List_of_Points) <1:													# Check there are the right number of inputs
		print2("Error - Check excel input Terminals there should be at least 1 Item in the list ")
	Load_Flow_Setting = analysis_dict["Loadflow_Settings"]						# Imports Settings for LDF calculation
	if len(Load_Flow_Setting) != 55:											# Check there are the right number of inputs
		print2('Error - Check excel input Loadflow_Settings there should be 55 Items in the list there are only: {} {}'
			   .format(len(Load_Flow_Setting), Load_Flow_Setting))
	Fsweep_Settings = analysis_dict["Frequency_Sweep"]							# Imports Settings for Frequency Sweep calculation
	if len(Fsweep_Settings) != 16:												# Check there are the right number of inputs
		print2('Error - Check excel input Frequency_Sweep there should be 16 Items in the list there are only: {} {}'
			   .format(len(Fsweep_Settings), Fsweep_Settings))
	Harmonic_Loadflow_Settings = analysis_dict["Harmonic_Loadflow"]				# Imports Settings for Harmonic LDF calculation
	if len(Harmonic_Loadflow_Settings) != 15:									# Check there are the right number of inputs
		print2('Error - Check excel input Harmonic_Loadflow Settings there should be 17 Items in the list there are only: {} {}'
			   .format(len(Harmonic_Loadflow_Settings), Harmonic_Loadflow_Settings))

	# Check all study cases converge, etc. and produce a new study case + operational scenario for each one
	# #List_of_Studycases1 = check_list_of_studycases(List_of_Studycases)			# This loops through all the studycases and operational scenarios listed and checks them skips any ones which don't solve
	# Adjusted to now return a list of handles to class <hast.pf.PF_Study_Case> which contain handles for the powerfactory
	# scenario objects that require activating.
	dict_of_projects = check_list_of_studycases(List_of_Studycases)

	# TODO: Find the points at which frequency sweeps are being run and instead create a ComFreqSweep and ComHLF within
	# TODO: each study case.
	# TODO: These can then be added to the ComTaskAutomation to run 3+ processes at a time.

	# Excel export contained within this loop
	# TODO: This section needs splitting up to firstly loop through creating the results folders and frequency sweeps,
	# TODO: these should be assigned to a task automation
	# TODO: Second stage is that these automated tasks are run
	# TODO: Third stage is that the results are retrieved
	# TODO: Comments have been put in below with the relevant locations which need looking into
	if FS_Sim or HRM_Sim:
		# TODO: Create task automation command and store it in general studies folder (ComTasks)

		FS_Contingency_Results, HRM_Contingency_Results = [], []
		count_studycase = 0

		# List of projects are created and then a unique list is used to iterate through for running studies in parallel
		prj_list = []

		# Get and deactivate current project
		current_prj = app.GetActiveProject()
		current_prj.Deactivate()

		# #while count_studycase < len(List_of_Studycases1):												# Loop Through (Study Cases, Operational Scenarios)
		t1 = time.clock()
		for prj_name, prj_cls in dict_of_projects.items():
			# Activate project
			prj_activation_failed = prj_cls.prj.Activate()

			#If failed activation then returns 1 (i.e. True) and for loop is continued
			if prj_activation_failed == 1:
				print2('Not possible to activate project {} and so no studies are performed for this project'
					   .format(prj_cls.name))
				continue

			t1_prj_start = time.clock()
			print1('Creating studies for Study Cases associated with project {}'.format(prj_cls.name))

			for count_studycase, study_cls in enumerate(prj_cls.sc_cases):  # Loop Through (Study Cases, Operational Scenarios)
				# Results folder created for each project
				# TODO: Check project isn't already active before activating
				# #prj = study_cls.prj.Activate()  # Activate Project
				print1('Creating studies for study case {} with operational scenario {} and contingency {}'
					   .format(study_cls.sc_name, study_cls.op_name, study_cls.cont_name))

				# #MOVED - prj = activate_project(List_of_Studycases1[count_studycase][1])								# Activate Project
				# #if len(str(prj)) > 0:

				# TODO: Need to add back in error checking
				# #StudyCase, study_error = activate_study_case(List_of_Studycases1[count_studycase][2])		# Activate Study Case in List
				# Activate Study Case
				StudyCase = study_cls.sc
				StudyCase.Activate()
				# TODO: This could be done as part of class when initialised
				# Add study case to task automation
				study_cls.task_auto.AppendStudyCase(study_cls.sc)

				# #Scenario, scen_error = activate_scenario(List_of_Studycases1[count_studycase][3])		# Activate corresponding operational Scenario
				# Activate Scenario
				Scenario = study_cls.op
				Scenario.Activate()
				Study_Case_Folder = app.GetProjectFolder("study")										# Returns string the location of the project folder for "study", (Ops) "scen" , "scheme" (Variations) Python reference guide 4.6.19 IntPrjfolder
				Operation_Case_Folder = app.GetProjectFolder("scen")
				# #Variations_Folder = app.GetProjectFolder("scheme")
				# TODO: Clarify purpose of variation
				# #Active_variations = get_active_variations()
				# #Variation = create_variation(Variations_Folder,"IntScheme",Variation_Name)
				# #activate_variation(Variation)
				# #Stage = create_stage(Variation,"IntSstage",Variation_Name)
				# #New_Contingency_List, Con_ok = check_contingencies(List_of_Contingencies) 				# Checks to see all the elements in the contingency list are in the case file
				Terminals_index, Term_ok = check_terminals(List_of_Points)								# Checks to see if all the terminals are in the case file skips any that aren't

				# TODO: Results folder should now be declared above rather than being related to a particular
				# TODO study case.
				# #studycase_results_folder, folder_exists1 = create_folder(StudyCase, Results_Folder)
				# #op_sc_results_folder, folder_exists2 = create_folder(Operation_Case_Folder, Operation_Scenario_Folder)
				Net_Elm1 = get_object(Net_Elm)															# Gets the Network Elements ElmNet folder
				if len(Net_Elm1) < 1:
					print2("Could not find Network Element folder, Note: this is case sensitive :" + str(Net_Elm))
				if len(Terminals_index) > 1 and Excel_Export_Z12:
					# TODO: Move mutual folder to central study case results folder
					studycase_mutual_folder, folder_exists3 = create_folder(Net_Elm1[0], Mut_Elm_Fld)		# Create Folder for Mutual Elements
					List_of_Mutual = create_mutual_impedance_list(studycase_mutual_folder, Terminals_index)	# Create List of mutual impedances between the terminals in the folder
				else:
					Excel_Export_Z12 = False															# Can't Export mutual impedances if you give it only one bus
					List_of_Mutual = []
					studycase_mutual_folder = ''
				# #count = 0

				# No need to loop through contingencies any more since each study case is contingency specific
				# #while count < len(New_Contingency_List):												# Loop Through Contingency list
				# #	print1("Carrying out Contingency: " + New_Contingency_List[count][0], bf=2, af=0)
				# # deactivate_scenario()																# Can't copy activated Scenario so deactivate it

				# #	object_exists, new_object = check_if_object_exists(op_sc_results_folder, List_of_Studycases1[count_studycase][0] + str("_" + New_Contingency_List[count][0] + ".IntScenario"))
				# #	if object_exists == 0:
				# #		new_scenario = add_copy(op_sc_results_folder,Scenario,List_of_Studycases1[count_studycase][0] + str("_" + New_Contingency_List[count][0]))	# Copies the base scenario
				# #	else:
				# #		new_scenario = new_object[0]
				# #	scen_error = activate_scenario1(new_scenario)										# Activates the base scenario
				# #	if New_Contingency_List[count][0] != "Base_Case":									# Apply Contingencies if it is not the base case
				# #		for switch in New_Contingency_List[count][1:]:
				# #			switch_coup(switch[0],switch[1])
				# #	save_active_scenario()

				# TODO: create results file and associated command, save results to results folder and command to study case
				if FS_Sim:
					# #sweep = create_results_file(studycase_results_folder, New_Contingency_List[count][0] + "_FS",9)		# Create Results File
					sweep = create_results_file(study_cls.res_folder, study_cls.name + "_FS", 9)  # Create Results File
					trm_count = 0
					while trm_count < len(Terminals_index):											# Add terminal variables to the Results file
						add_vars_res(sweep, Terminals_index[trm_count][3], FS_Terminal_Variables)
						trm_count = trm_count + 1
					if Excel_Export_Z12:
						for mut in List_of_Mutual:													# Adds the mutual impedance data to Results File
							add_vars_res(sweep, mut[2], Mutual_Variables)

					# Create command for frequency sweep and add to Task Automation
					freq_sweep = study_cls.create_freq_sweep(results_file=sweep, settings=Fsweep_Settings)
					# Add freq_sweep to task automation
					study_cls.task_auto.AppendCommand(freq_sweep, 0)
					print1('Frequency sweep added for study case {}'.format(study_cls.name))

					# #Fsweep_err_cde = freq_sweep(sweep, Fsweep_Settings)								# Carry out Frequency Sweep


					# #TODO: results processing needs to happen later once parallel processing has been completed
					# #if Fsweep_err_cde == 0:															# Skips the contingency if Frequency Sweep doesn't solve
					# #	fs_scale, fs_res = retrieve_results(sweep,0)
					# #	fs_scale.insert(1,"Frequency in Hz")										# Arranges the Frequency Scale
					# #	fs_scale.insert(1,"Scale")
					# #	fs_scale.pop(3)
					# #	for tope in fs_res:															# Adds the additional information to the results file
					# #		tope.insert(1,New_Contingency_List[count][0])							# Op scenario
					# #		tope.insert(1,List_of_Studycases1[count_studycase][0])					# Study case description
					# #		FS_Contingency_Results.append(tope)										# Results
					# #else:
					# #	print2("Error with Frequency Sweep Simulation: " + List_of_Studycases1[count_studycase][0] + New_Contingency_List[count][0])
				else:
					# Frequency sweep not carried out so no need to add to task automation
					print1('No frequency sweep included for study case {}'.format(study_cls.name))
					# #fs_scale = []
				if HRM_Sim:
					# Create a results file to store the results from the harmonic load flow
					harm = create_results_file(study_cls.res_folder, study_cls.name + "_HLF", 6)		# Creates the Harmonic Results File
					# #harm = create_results_file(studycase_results_folder, New_Contingency_List[count][0] + "_HLF",6)		# Creates the Harmonic Results File
					trm_count = 0
					while trm_count < len(Terminals_index):											# Add terminal variables to the Results file
						add_vars_res(harm, Terminals_index[trm_count][3], HRM_Terminal_Variables)
						trm_count = trm_count + 1

					# Create command for harmonic load flow and add to Task Automation
					hldf_command = study_cls.create_harm_load_flow(results_file=harm,
																   settings=Harmonic_Loadflow_Settings)
					study_cls.task_auto.AppendCommand(hldf_command, 0)
					print1('Harmonic load flow added for study case {}'.format(study_cls.name))
					# #TODO: Results processing needs to happen elsewhere
					# #Harm_err_cde = harm_load_flow(harm, Harmonic_Loadflow_Settings)
					# #if Harm_err_cde == 0:
					# #	hrm_scale, hrm_res = retrieve_results(harm,1)
					# #	hrm_scale.insert(1,"THD")													# Inserts the THD
					# #	hrm_scale.insert(1,"Harmonic")												# Arranges the Harmonic Scale
					# #	hrm_scale.insert(1,"Scale")
					# #	hrm_scale.pop(4)															# Takes out the 50 Hz
					# #	hrm_scale.pop(4)
					# #	for res12 in hrm_res:
					# #		thd1 = re.split(r'[\\.]',res12[1])
					# #		thd2 = app.GetCalcRelevantObjects(thd1[11] + ".ElmSubstat")
					# #		thdz = False
					# #		if thd2[0] is not None:
					# #			thd3 = thd2[0].GetContents()
					# #			for thd4 in thd3:
					# #				if (thd1[13] + ".ElmTerm") in str(thd4):
					# #					THD = thd4.GetAttribute('m:THD')
					# #					thdz = True
					# #		elif thd2[0] is not None or thdz == False:
					# #			THD = "NA"
					# #		#thd4 = app.SearchObjectByForeignKey(thd1[11] + ".ElmSubstat")
					# #		# TODO: THD recognised as undeclared which will be solved by wrapping into function
					# #		res12.insert(2, THD)														# Insert THD
					# #		res12.insert(2, New_Contingency_List[count][0])							# Op scenario
					# #		res12.insert(2, List_of_Studycases1[count_studycase][0])					# Study case description
					# #		res12.pop(5)
					# #		HRM_Contingency_Results.append(res12)									# Results
					# #else:
					# #	print2("Error with Harmonic Simulation: " + List_of_Studycases1[count_studycase][0] + New_Contingency_List[count][0])

				else:
					print1('No Harmonic load flow added for study case {}'.format(study_cls.name))
					# #hrm_scale =[]
				# #count = count + 1
				# #print1("", bf=2, af=0)

				# TODO:  Need to add in statement that will mean the initial study case is restored once completed
				# #Scenario, scen_error = activate_scenario(List_of_Studycases1[count_studycase][3])		# Activate the base case scenario this just ensures that when the script finishes using PF that it goes back to a regular case
				# #count_studycase = count_studycase + 1
				# #print1("", bf=2, af=0)
				# TODO: Delete folders once created, if requested to do so in settings
				# #if Delete_Created_Folders:
				# #	delete_object(studycase_results_folder)
				# #	delete_object(op_sc_results_folder)
				# #	if Excel_Export_Z12:
				# #		delete_object(studycase_mutual_folder)
				# #	# Deactivate active variation
				# #	Variation.Deactivate()
				# #	delete_object(Variation)
				# #else:
				# #	print2('Could Not Activate Project {}'.format(study_cls.prj))
				# #	# #print2('Could Not Activate Project: {}'.format(List_of_Studycases1[count_studycase][1]))
				# #	# ERROR:  print2("Could Not Activate Project: " + Project_Name)

			print1('Creating of commands for studies in project {} completed in {:0.2f} seconds'
				   .format(prj_cls.name, time.clock()-t1_prj_start))
			t1_prj_studies = time.clock()

			print1('Parallel running of frequency scans and harmonic load flows associated with project {}'
				.format(prj_cls.name))

			# Call Task automation to run studies
			prj_cls.task_auto.Execute()

			print1('Studies for project {} completed in {:0.2f} seconds'
				   .format(prj_cls.name, time.clock()-t1_prj_studies))

			# Once studies complete, deactivate project
			prj_cls.prj.Deactivate()

		print1('PowerFactory studies all completed in {:0.2f} seconds'.format(time.clock()-t1))

		# TODO:  At this point the task automation files have been created and will now need to loop through each project
		# TODO: iteratively to run the commands before then processing the results

		print2(' DEBUG FORCED EXIT ')
		raise SyntaxError('EXIT')

		if Export_to_Excel:																# This Exports the Results files to Excel in terminal format
			print1("\nProcessing Results and output to Excel", bf=1, af=0)
			start2 = time.clock()																# Used to calc the total excel export time
			# Create a new instance of excel to deal with reading and writing of data to excel instance
			# With statement means that even if error occurs new instance of excel is closed
			with hast.excel_writing.Excel(print_info=print1, print_error=print2) as excel_cls:
				wb = excel_cls.create_workbook(workbookname=Excel_Results, excel_visible=Excel_Visible)	# Creates Workbook
				trm1_count = 0
				while trm1_count < len(Terminals_index):											# For Terminals in the index loop through creating results to pass to excel sheet creator
					start3 = time.clock()															# Used for measuring time to create a sheet
					FS_Terminal_Results = []														# Creates a Temporary list to pass through terminal data to excel to create the terminal sheet
					if FS_Sim:
						start4 = time.clock()
						FS_Terminal_Results.append(fs_scale)										# Adds the scale to terminal
						for results34 in FS_Contingency_Results:									# Adds each contingency to the terminal results
							if str(Terminals_index[trm1_count][3]) == results34[3]:					# Checks it it the right terminal and adds it
								results34.pop(3)													# Takes out the terminal  PF object (big long string)
								FS_Terminal_Results.append(results34)								# Append terminal data to the results list to be later passed to excel
						#print1(1,"Process Results RX & Z in Python: " + str(round((time.clock() - start4),2)) + " Seconds",0)		# Returns python results processing time
						if Excel_Export_Z12:
							start5 = time.clock()
							for results35 in FS_Contingency_Results:								# Adds each contingency to the terminal results
								for tgb in List_of_Mutual:
									if Terminals_index[trm1_count][3] == tgb[3]:
										if str(tgb[2]) == str(results35[3]):						# Checks it it the right terminal and adds it
											results35.pop(3)										# Takes out the terminal  PF object (big long string)
											results35.insert(0,tgb[1])								# Adds in the Mutual tag ie Letterkenny_Binbane
											FS_Terminal_Results.append(results35)					# If it is the right terminal append
							print1("Process Results Z12 in Python: " + str(round((time.clock() - start5),2)) + " Seconds",
								   bf=1, af=0)		# Returns python results processing time
					HRM_Terminal_Results = []														# Creates a Temporary list to pass through terminal data to excel to create the terminal sheet
					if HRM_Sim:
						start6 = time.clock()
						# TODO: Error reported that fs_scale can be undefined.  Wrapping in class / function will prevent this
						HRM_Terminal_Results.append(hrm_scale)										# Adds the scale to terminal
						if Excel_Export_HRM:
							for results35 in HRM_Contingency_Results:								# Adds each contingency to the terminal results
								if str(Terminals_index[trm1_count][3]) == results35[1]:				# Checks it it the right terminal and adds it
									results35.pop(1)												# Takes out the terminal  PF object (big long string)
									HRM_Terminal_Results.append(results35)							# Append terminal data to the results list to be later passed to excel
						print1("Process Results HRM in Python: " + str(round((time.clock() - start6),2)) + " Seconds",
							   bf=1, af=0)		# Returns python results processing time

					# Replaced with using instance in excel_writing
					excel_cls.create_sheet_plot(sheet_name=Terminals_index[trm1_count][0],
												fs_results=FS_Terminal_Results,
												hrm_results=HRM_Terminal_Results,
												wb=wb,
												# TODO:  The following are all booleans and could be passed in a better way
												excel_export_rx=Excel_Export_RX,
												excel_export_z=Excel_Export_Z,
												excel_export_hrm=Excel_Export_HRM,
												fs_sim=FS_Sim,
												excel_export_z12=Excel_Export_Z12,
												excel_convex_hull=Excel_Convex_Hull,
												hrm_sim=HRM_Sim)				# Uses the terminal results to create a sheet and graph
					# #create_sheet_plot(Terminals_index[trm1_count][0],FS_Terminal_Results, HRM_Terminal_Results, _wb=wb, _xl=xl)				# Uses the terminal results to create a sheet and graph
					trm1_count = trm1_count + 1
				# progress_txt = read_text_file(Progress_Log)

				# Save and close workbook
				excel_cls.close_workbook(wb=wb, workbookname=Excel_Results)
				# #close_workbook(_wb=wb, workbookname=Excel_Results, _xl=xl)																# Closes and saves the workbook
				print1("Total Excel Export Time: " + str(round((time.clock() - start2),2)) + " Seconds",
					   bf=1, af=0)	# Returns the Total Export time

	print1('Total Time: {:.2f}'.format(time.clock() - start),
		   bf=1, af=0)

	# Close the logger since script has now completed and this forces flushing of the open logs before script exits
	logger.close_logging()