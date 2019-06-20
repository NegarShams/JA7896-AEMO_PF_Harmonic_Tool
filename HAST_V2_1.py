"""
#######################################################################################################################
###											HAST_V2_0																###
###		Script initially produced by EirGrid for Harmonics Automated Simulation Tool and further developed by		###
###		David Mills (PSC) to improve performance, extracting of data to Excel and solve some errors present in 		###
###		the code.																									###
###		The script now makes use of PowerFactory parallel processing and will require that the Parallel Processing	###
###		function in PowerFactory has been enabled and the number of cores has been set to N-1						###
###																													###
###		Code layout has been updated to align with PEP																###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###		project JI6973 for EirGrid project PSPF010 - Specialise Support in Power Quality Analysis during 2018		###
###																													###
#######################################################################################################################

DEVELOPMENT CODE:  This code is still in development since has not been produced to account for use of Harmonic Load
Flow or extraction of ConvexHull data to excel.

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
- Converted to us a logging system that is stored in hast2.logger which will avoid writing to a log file every time
something happens
- Added functionality to repeat studies for different filter arrangements and will now also run in unattended mode
"""

DIG_PATH = r'C:\Program Files\DIgSILENT\PowerFactory 2016 SP5'
DIG_PYTHON_PATH = r'C:\Program Files\DIgSILENT\PowerFactory 2016 SP5\Python\3.5'

# IMPORT SOME PYTHON MODULES
# --------------------------------------------------------------------------------------------------------------------
import os
import sys
import importlib
import time
import shutil
import tkinter as tk
import distutils.version

# HAST module package requires reload during code development since python does not reload itself
# HAST module package used as functions start to be transferred for efficiency
import hast2_1
hast2 = importlib.reload(hast2_1)
import hast2_1.constants as constants
import hast2_1.pf as pf
import Process_HAST_extract

# Meta Data
__author__ = 'David Mills'
__version__ = '2.1.2'
__email__ = 'david.mills@PSCconsulting.com'
__phone__ = '+44 7899 984158'
__status__ = 'In Development - Beta'

# GLOBAL variable used to avoid trying to print to PowerFactory when running in unittest mode, set to true by unittest
DEBUG_MODE = False

hast_inputs_filename = 'HAST_Inputs.xlsx'

# Location of this script so can be used for any files that need to be located
filelocation = os.path.dirname(os.path.abspath(__file__))

# TODO:  Identify machine running so can adjust target folder appropriately.  May not be required and instead could
# TODO:  rename inputs file appropriately.

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
		logger.info('Activated Project Successfully: {}'.format(project))
		# prj renamed _prj to avoid shadowing name from parent project
		_prj = app.GetActiveProject()										# Get active project
	else:																	# Project Failed to Activate
		# Print Information to progress log and PowerFactory window and Error Log
		logger.error('Error Not able to Activate Project: {}....................'.format(project))
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
			logger.debug('Activated Study Case Successfully: {}'.format(study_case1[0]))
		else:
			logger.error('Error Unable to Activate Study Case: {}.............................'
						 .format(study_case))
	else:
		logger.error('Could not activate StudyCase as no matching name in case: {}'.format(study_case))
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
			logger.debug('Deactivated active study <{}> case successfully'.format(study))
		elif sce > 0:
			print2('Error Unsuccessfully Deactivated Study Case: {}..............................'.format(study))
			print2('Unsuccessfully Deactivated Scenario Error Code: {}'.format(sce))
	else:
		logger.debug("No Study Case active to deactivate ................................")
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
		logger.debug('Activated Scenario Successfully: {}'.format(scenario1[0]))
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
		logger.debug('Activated Scenario Successfully: {}'.format(scenario))
	elif sce == 1:
		logger.error('Error Unsuccessfully Activated Scenario: {}...............................'.format(scenario))
		logger.error('Unsuccessfully Activated Scenario Error Code: {}'.format(sce))
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
		elif sce > 0:
			logger.error('Error Unsuccessfully Deactivated Scenario: {}............'.format(scenario1))
			logger.error('Unsuccessfully Deactivated Scenario Error Code: {}'.format(sce))
	else:
		logger.debug('No Scenario Active to Deactivate ................................')
	return None


def save_active_scenario(): 		# Save active scenario
	"""
		Save the active scenario
	:return: None
	"""
	scenario1 = app.GetActiveScenario()
	sce = scenario1.Save()
	if sce==0:
		logger.debug('Saved active scenario successfully: {}'.format(scenario1))
	elif sce == 1 and scenario1 is None:
		logger.error('Error unsuccessfully saved scenario: {}'.format(scenario1))
		logger.error('Unsuccessfully saved scenario error code: {}'.format(sce))
	else:
		logger.debug('No Scenario Active to Save.........................................')
	return None


def get_active_variations():			# Get Active Network Variations
	"""
		Get active variations
	:return list variations: Returns list of variations currently active
	"""
	variations =  app.GetActiveNetworkVariations()
	logger.info('Current Active Variations: ')
	if len(variations) > 1:
		for item in variations:
			aa = str(item)
			pp = aa.split("Variations.IntPrjfolder\\")
			ss = pp[1]
			tt = ss.split(".IntScheme")
			logger.info('\t{}'.format(tt[0]))
	elif len(variations) == 1:
		logger.info(variations)
	else:
		logger.info('No Variations Active')
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
	logger.debug('Variation {} created'.format(variation))
	return variation


def activate_variation(variation): 		# Activate Scenario
	""" 
		Activate previously created variation
	:param powerfactory.Variation variation: handle to existing powerfactory variation
	:return int sce: Integer (0,1) on whether success or fail on activating variation
	"""
	sce = variation.Activate() 											# Activate Study case
	if sce == 0:
		logger.debug('Activated Variation Successfully: {}'.format(variation))
	elif sce == 1:
		logger.error('Error Unsuccessfully Activated Variation: {}........................'.format(variation))
		logger.error('Unsuccessfully Activated Variation Error Code: {}'.format(sce))
	return sce


def create_stage(location, pfclass, name):
	"""
		Creates a new expansion stage in powerfactory
	:param powerfactory.Location location: Handle to powerfacory location
	:param str pfclass: String describing the powerfactory stage to be created
	:param ztr name: Name of new stage to be created
	:return powerfactory.Stage stage: Handle to newly created powerfactory stage
	"""
	stage = create_object(location=location,
						  pfclass=pfclass,
						  name=name)
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
		logger.debug('Activated Variation Stage Successfully: {}'.format(stage))
	elif sce != 0:
		logger.error('Unable to activate variation expansion stage: {} and PowerFactory returned error code: {}'
					 .format(stage,sce))
	return None


def load_flow(load_flow_settings, sc, studycase_name=''):		# Inputs load flow settings and executes load flow
	"""
		Run load flow in powerfactory
	:param list load_flow_settings: List of settings for powerfactory when running loadflow
	:param sc:  Studycase handle
	:param str studycase_name:  Name of study case being run to include in error message reporting	
	:return (int error_code, ldf): Error code provided by powerfactory determining its success, 
									handle to powerfactory load flow command 
	"""
	# TODO: Setting should only need setting once rather than every time load_flow is run so could be defined in
	# TODO: + constants
	t1 = time.clock()
	## Loadflow settings
	## -------------------------------------------------------------------------------------
	# Create new object for the load flow on the base case so that existing settings are not overwritten
	ldf = create_object(location=sc,
						pfclass=constants.PowerFactory.ldf_command,
						name=constants.PowerFactory.default_ldf_name)
	# Get handle for load flow command from study case
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

	error_code = ldf.Execute()
	t2 = time.clock() - t1
	if error_code == 0:
		logger.info('\t - Load Flow calculation successful for {}, time taken: {:.2f} seconds'
					.format(studycase_name, t2))
	elif error_code == 1:
		logger.error('Load Flow for {} failed due to divergence of inner loops, time taken: {:.2f} seconds'
					 .format(studycase_name, t2))
	elif error_code == 2:
		logger.error('Load Flow failed for {} due to divergence of outer loops, time taken: {:.2f} seconds'
					 .format(studycase_name, t2))
	return error_code, ldf

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
	# Get handle for harmonic load flow command from study case
	hldf = app.GetFromStudyCase(constants.PowerFactory.hldf_command)

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
		logger.debug('Harmonic Load Flow calculation successful: {:.2f} seconds'.format(t2))
	elif error_code > 0:
		logger.error('Harmonic Load Flow calculation unsuccessful: {:.2f} seconds.............'.format(t2))
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
	# Get handle for harmonic load flow command from study case
	frq = app.GetFromStudyCase(constants.PowerFactory.frq_sweep_command)
	# Basic
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
		logger.debug('Frequency Sweep calculation successful, time taken: {:.2f} seconds'.format(t2))
	elif error_code > 0:
		logger.error('Frequency Sweep calculation unsuccessful, time taken: {:.2f} seconds.......'.format(t2))
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
		logger.debug('Switching Element: {} Out of service "'.format(element))
	if service == 1:
		logger.debug('Switching Element: {} In to service '.format(element))
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
		logger.debug('Folder already exists: {}'.format(name))
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
	logger.debug('Creating Folder: {}'.format(name))
	folder1, folder_exists = check_if_folder_exists(location, name)				# Check if the Results folder exists if it doesn't create it using date and time
	logger.debug('Location = {}'.format(location))
	logger.debug('Folder1 = {}'.format(folder1))
	logger.debug('Folder exists = {}'.format(folder_exists))
	if folder_exists == 0:
		logger.debug('Creating new folder of type = {} with name {}'
					 .format(constants.PowerFactory.pf_folder_type, name))
		_new_object = location.CreateObject(constants.PowerFactory.pf_folder_type, name)
		logger.debug('Newly created folder = {}'.format(_new_object))
		# loc_name = name							# Name of Folder
		# owner = "Barry"							# Owner
		# iopt_sys = 0							# Attributes System
		# iopt_type = 0							# Folder Type 0 Common
		# for_name = ""							# Foreign key
		# desc = ""								# Description
	else:
		_new_object = folder1[0]
	logger.debug('New folder created = {}'.format(_new_object))
	return _new_object, folder_exists

# Creates a mutual Impedance list from the terminal list in a folder under the active studycase
def create_mutual_impedance_list(location, terminal_list, list_of_mutual = [], list_of_mutual_names = []):
	"""
		Create a mutual impedance list from the terminal list in a folder under the active studycase
	:param powerfacory.Location location: Handle for location to be created 
	:param list terminal_list: List of terminals
	:param list list_of_mutual:  (optional=[]) List of mutual impedance values already in list
	:param list list_of_mutual_names:  List of names for mutual impedance values already created
		(these are included in both directions 'from_to', 'to_from' to avoid duplication)
	:return list list_of_mutual: List of mutual impedances
	"""
	logger.info('Creating: Mutual Impedance List of Terminals')
	terminal_list1 = list(terminal_list)

	# Produce a dictionary of the terminals so can lookup which ones require mutual impedance data
	_dict_terminal_mutual = {term[3]:term[4] for term in terminal_list1}
	# If some mutual elements have already been created then those will not be created again
	# list_of_mutual = existing_list_of_mutual
	# list_of_mutual_names = []
	# TODO: Improve since this is a loop of loops
	for _y in terminal_list1:
		pf_terminal_1 = _y[3]
		for _x in terminal_list1:
			pf_terminal_2 = _x[3]
			# Adjusted so that mutual data will only be collected from this node to the remote node if the remote node
			# is set to True in the input data (column 4)
			if pf_terminal_2 != pf_terminal_1 and \
					(_dict_terminal_mutual[pf_terminal_2] or
					 _dict_terminal_mutual[pf_terminal_1]):
				name = '{}_{}'.format(_y[0],_x[0])
				# Inverse name created for checking if already in list
				name_inverse = '{}_{}'.format(_x[0], _y[0])
				# Checks that mutual impedance has not already been created for the reverse direction
				if name not in list_of_mutual_names:
					logger.debug('Term 1 = {} - {}, Term 2 = {} - {}'.format(pf_terminal_1, _y, pf_terminal_2, _x))
					elmmut = create_mutual_elm(location, name, pf_terminal_1, pf_terminal_2)
					list_of_mutual.append([str(_y[0]), name, elmmut, pf_terminal_1, pf_terminal_2])

					# Add name in both directions to list_of_mutual created
					list_of_mutual_names.append(name)
					list_of_mutual_names.append(name_inverse)
				else:
					logger.debug('Mutual elements {} has already been created in the form {}'
								 .format(name, name_inverse))
	return list_of_mutual, list_of_mutual_names


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
		logger.debug('Object Successfully Deleted: {}'.format(object_to_delete))
	else:
		logger.info('Error deleting object: {}.....................'.format(object_to_delete))
	return None


def check_if_object_exists(location, name):  	# Check if the object exists
	"""
		Check if an object exists by name
	:param powerfactory.Location location: Handle for PF location to investigate 
	:param str name: Name of object to look for 
	:return object_exists, new_object: True / False on whether object exists, handle for powerfactory.Object 
	"""
	logger.debug('{} {}'.format(location, name))
	#_new_object used instead of new_object to avoid shadowing
	_new_object = location.GetContents(name)
	object_exists = 0
	if len(_new_object) > 0:
		logger.debug('Object Exists: {}'.format(name))
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
		logger.debug('Copying object {} successful'.format(object))
	else:
		logger.error('Error AddCopy Unsuccessful: {} to {} as {}'.format(object, folder, name1))
	return new_object


def create_object(location, pfclass, name):			# Creates a database object in a specified location of a specified class
	"""
		Creates a database object in a specified location of a specified class
	:param powerfactory.Location location: Location in which new object should be created 
	:param str pfclass: Type of element to be created 
	:param str name: Name to be given to new object 
	:return powerfactory.Object _new_object: Handle to object returns 
	"""
	# Checks if object already exists before creating a new one and return handle to object if already
	# exists.  Returns list of all objects that match the name provided
	existing_object = location.GetContents('{}.{}'.format(name,pfclass))
	if existing_object:
		_new_object = existing_object[0]
	else:
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
	sweep.header = ["Results File"]
	sweep.desc = ["Results type {}".format(type_of_file)]
	return sweep


def create_study_case_results_files(cls_sc, cls_prj):
	"""

	:param pf_Study_Case sc:  Handle to the power factory study case (must be active)
	:param str sc_name:  Handle to the power factory study case (must be active)
	:return:
	"""

	_t1_prj_start = time.clock()

	# Determine the terminals and mutual impedance data requested and check they exist within this project case

	logger.info('Checking all terminals and producing mutual impedance data')
	# Checks to see if all the terminals are in the project and skips any that aren't
	if cls_prj.terminals_index is None:
		cls_prj.terminals_index, _ = check_terminals(List_of_Points, prj_name=cls_prj.name, sc_name=cls_sc.name)
		# Gets the network elements folder
		list_network_element_folders = get_object(Net_Elm)
		logger.info('Network elements folder = {}'.format(list_network_element_folders))

		if len(list_network_element_folders) < 1:
			logger.error('Could not find Network Element folder, Note: this is case sensitive : {}'.format(Net_Elm))
		else:
			cls_prj.folder_network_elements = list_network_element_folders[0]

	# Add mutual_impedance links to the study folders
	if len(cls_prj.terminals_index) > 1 and cls_prj.include_mutual:
		# Results for mutual impedance have to be stored in the Network Folder
		# Create folder for mutual elements
		logger.debug('Creating mutual impedance data being added')
		logger.info('Mutual impedance folder = {}'.format(cls_prj.mutual_impedance_folder))
		active_project = app.GetActiveProject()
		logger.info('Active project = {}'.format(active_project))
		if cls_prj.mutual_impedance_folder is None:
			# Initial folder is created in the Project network data folder and is then moved to the EirGrid network
			# elements folder.  This is required to resolve issues when running in unattended mode.
			network_data_folder = app.GetProjectFolder(constants.PowerFactory.pf_prjfolder_type)
			# Create mutual impedance folder
			cls_prj.mutual_impedance_folder, folder_exists3 = create_folder(network_data_folder, Mut_Elm_Fld)
			logger.debug('New mutual impedance folder = {}'.format(cls_prj.mutual_impedance_folder))
			# Move newly created folder to ElmNet
			net_elm = get_object(Net_Elm)[0]
			failed = cls_prj.folder_network_elements.Move(cls_prj.mutual_impedance_folder)
			if failed == 0:
				logger.debug('Moving mutual impedance folder {} from {} to {} was a success'
							.format(cls_prj.mutual_impedance_folder, network_data_folder,
									cls_prj.folder_network_elements))
			else:
				logger.error(('Moving mutual impedance folder {} from {} to {} failed and so no mutual '
							  'impedance data will be exported')
							 .format(cls_prj.mutual_impedance_folder, network_data_folder,
									 cls_prj.folder_network_elements))
				cls_prj.include_mutual = False

		# Newly created folder is added to list of folders created so can be deleted at end of study
		# No longer required since Variation deleted which includes the mutual folder
		# Create list of mutual impedances between the terminals in the folder requested
		cls_prj.list_of_mutual, cls_prj.list_of_mutual_names = create_mutual_impedance_list(
			location=cls_prj.mutual_impedance_folder,
			terminal_list=cls_prj.terminals_index,
			list_of_mutual=cls_prj.list_of_mutual,
			list_of_mutual_names=cls_prj.list_of_mutual_names
		)
		msg1 = 'The following mutual impedance elements were created and will be monitored:'
		msg2 = '\n'.join(['\t - Name: {} \n\t\t PowerFactory Element: {}'.format(x[1], x[2]) for x in cls_prj.list_of_mutual])
		logger.info('{}\n{}'.format(msg1,msg2))
	else:
		# Can't Export mutual impedances if you give it only one bus
		cls_prj.include_mutual = False
		logger.warning('Unable to create mutual impedance terminals')

	if FS_Sim:
		# During task automation each process only has access to single study case and therefore results
		# need to be stored in the study case file.  Once completed they can then be moved to a centralised
		# results folder
		cls_sc.fs_res = create_results_file(cls_sc.sc,
											constants.PowerFactory.default_results_name + constants.PowerFactory.default_fs_extension,
											9)  # Create Results File
		for term in cls_prj.terminals_index:
			add_vars_res(cls_sc.fs_res, term[3], constants.HASTInputs.fs_term_variables)

		if cls_prj.include_mutual:
			for mut in cls_prj.list_of_mutual:								# Adds the mutual impedance data to Results File
				add_vars_res(cls_sc.fs_res, mut[2], constants.HASTInputs.mutual_variables)
		cls_sc.fs_res.SetAsDefault()
	else:
		# Frequency sweep not carried out so no need to add to task automation
		logger.debug('No frequency sweep included for study case {}'.format(cls_sc.name))

	if HRM_Sim:
		# During task automation each process only has access to single study case and therefore results
		# need to be stored in the study case file.  Once completed they can then be moved to a centralised
		# results folder
		cls_sc.hldf_results = create_results_file(cls_sc.sc,
												  constants.PowerFactory.default_results_name + constants.PowerFactory.default_hldf_extension,
												  6)		# Creates the Harmonic Results File
		for term in cls_prj.terminals_index:									# Add terminal variables to the Results file
			add_vars_res(cls_sc.hldf_results, term[3], constants.HASTInputs.hldf_term_variables)

		cls_sc.hldf_results.SetAsDefault()

	else:
		logger.debug('No Harmonic load flow added for study case {}'.format(cls_sc.name))

	logger.info('Creating of commands for studies in project {} completed in {:0.2f} seconds'
				.format(cls_prj.name, time.clock() - _t1_prj_start))
	return None


def copy_study_case(name, sc_target_folder, sc, op_target_folder, op):
	"""
		Copy an existing study case including operational scenario and return reference to newly created study case
	:param str name: Name for new study cases and operational scenarios
	:param sc_target_folder:  Folder for new study case to be copied in
	:param sc: study case to be copied
	:param op_target_folder: Folder for new operational scenario to be copied into
	:param op: operational scenario to be copied
	:return (new_sc, new_op):
	"""

	# Deactivate study case so that it can be copied
	deactivate_study_case()

	# Copy study case
	new_sc = add_copy(sc_target_folder,
							   sc,
							   name)
	# Copy scenario
	new_op = add_copy(op_target_folder, op,
							 name)

	# Activate new study case and scenario
	new_sc.Activate()
	_ = activate_scenario1(new_op)


	return new_sc, new_op

def add_all_filters(new_filter_list, cont_name, sc, op, sc_target_folder,
					op_target_folder, variations_folder, list_params, cont_short_name,
					cls_prj, create_studies=False):
	"""
		Function to create new copy of study case for each filter
	:param new_filter_list:
	:param str cont_name:
	:param sc:  Handle to the study case in PowerFactory that will be copied
	:param op:  Handle the operational scenario in PowerFactory that will be copied for the new filter
	:param sc_target_folder:  Target folder in PowerFactory to contain the new study cases with the filters
	:param op_target_folder:  Target folder in PowerFactory to contain the new operational scenario
	:param variations_folder:  Folder that contains the variations
	:param list_params: Parameters from HAST inputs for the study case (i.e. name, project, etc.)
	:param cont_short_name:  Short name for the contingency under consideration
	:param pf.PFProject cls_prj: Handle to the custom PowerFactory project class that will house the study case
	:param bool create_studies: Whether to create the frequency scan and harmonic load flow studies for the filter case
	:return:
	"""
	for filter_item in new_filter_list:
		logger.debug('Adding filters under name {} to model'.format(filter_item.name))

		# Loop through each of the q_f values for this filter and add to PF
		for f_q in filter_item.f_q_values:
			# Create name for filter based on contingency, frequency, mvar value
			filter_name = '{}_{:.1f}Hz_{:.1f}MVAR'.format(filter_item.name, f_q[0], f_q[1])
			filter_full_name = '{}_{}'.format(cont_name, filter_name)

			filter_study_case, filter_op = copy_study_case(name=filter_full_name,
									 sc_target_folder=sc_target_folder,
									 sc=sc,
									 op_target_folder=op_target_folder,
									 op=op)

			logger.debug('New study case created and activated for modelling filter: {}'
						 .format(filter_full_name))

			# Create new variation specifically for this filter so can deactivate
			# before copying to ensure filter isn't added to every case
			filter_variation = create_variation(
				variations_folder,
				constants.PowerFactory.pf_scheme,
				filter_full_name)
			activate_variation(filter_variation)
			# Create and activate recording stage within variation
			_ = create_stage(filter_variation,
							 constants.PowerFactory.pf_stage,
							 filter_full_name)
			logger.debug('New variation created for filter: {}'.format(filter_full_name))

			# Add filter to model
			pf.add_filter_to_pf(
				_app=app,
				filter_name=filter_full_name,
				filter_ref=filter_item,
				q=f_q[1], freq=f_q[0],
				logger=logger)

			# Save updated scenario which now includes filter
			save_active_scenario()

			# Create new class reference with all the details for this contingency and
			# filter combination and then add to
			# list to be returned
			_study_cls = pf.PFStudyCase(full_name=filter_full_name,
										list_parameters=list_params,
										cont_name=cont_short_name,
										filter_name=filter_name,
										sc=filter_study_case,
										op=filter_op,
										prj=cls_prj,
										task_auto=cls_prj.task_auto,
										uid=start1,
										results_pth=Temp_Results_Export)
			_study_cls.create_load_flow(load_flow_settings=Load_Flow_Setting)

			# Determine if load flow successful and if not then don't include _study_cls in results
			if _study_cls.run_load_flow():

				# Create the frequency sweep and harmonic load flow studies
				if create_studies:
					_study_cls.create_studies(logger=logger,
											  fs=FS_Sim, hldf=HRM_Sim,
											  fs_settings=Fsweep_Settings,
											  hldf_settings=Harmonic_Loadflow_Settings)

				# Add study case to file
				cls_prj.sc_cases.append(_study_cls)
				logger.debug('Filter added to model and load flow run successfully for {}'
							 .format(filter_study_case))
			else:
				logger.error(
					('Load flow for filter study case {} not successful and so frequency scans ' +
					 ' and harmonic load flows will not be run for this case')
						.format(filter_study_case))

	return None


def check_list_of_studycases(list_to_check):		# Check List of Projects, Study Cases, Operational Scenarios,
	"""
		Check list of projects, study cases, operational scenarios, etc. solve for load flow
		Produces a new operational scenario and study case of each contingency so that each study case can be split
		out into separate parallel processing functions
	:param list list_to_check: List of items to check 
	:return dict prj_dict:  Dictionary of projects where the study cases associated with each project are contained within 
	"""
	time_sc_check = time.clock()
	logger.info(('Checking all Projects, Study Cases and Scenarios Solve for Load Flow, it will also check N-1 and ' 
				 ' create the operational scenarios if Pre_Case_Check is True\n'))
	new_list =[]

	# Empty list which will be populated with the new classes
	prj_dict = dict()
	# while _count_studycase < len(list_to_check):
	# TODO: Handling non-unique studycases
	for sc_list_parameters in list_to_check:
		logger.info('----####---- \t Studies for {} \t ----####----'.format(sc_list_parameters[0]))
		# TODO: Efficiency - This is activating a new project even if it is the same
		# TODO:  Create frequency / HLF command and results at this point and then copy / paste
		project_name = sc_list_parameters[1]
		_prj = activate_project(project_name)  # Activate Project

		if len(str(_prj)) > 0:
			study_case, _study_error = activate_study_case(sc_list_parameters[2])									# Activate Case
			if _study_error == 0:
				scenario, scen_err = activate_scenario(sc_list_parameters[3])										# Activate Scenario
				if scen_err == 0:
					logger.info('Load flow being run for HAST study case {}'.format(sc_list_parameters))
					ldf_err, ldf_command = load_flow(load_flow_settings=Load_Flow_Setting,
													 sc=study_case,
													 studycase_name=sc_list_parameters)																			# Perform Load Flow
					logger.debug('Load flow study completed with error code {}'.format(ldf_err))
					# Not possible to skip unsolved load flows since gets stuck in a loop
					if ldf_err == 0:
						# If no error then deletes the load flow command that was created to avoid leaving mess in
						# studycase
						delete_object(ldf_command)

						new_list.append(sc_list_parameters)
						logger.debug('Studycase Scenario Solving added to analysis list {}'
									 .format(sc_list_parameters))

						new_contingency_list, con_ok = check_contingencies(List_of_Contingencies) 				# Checks to see all the elements in the contingency list are in the case file
						# Adjusted to create new study_case for each op_scenario
						new_filter_list, filter_ok = check_filters(List_of_Filters)


						study_case_folder = app.GetProjectFolder('study')
						study_case_results_folder, _folder_exists2 = create_folder(study_case_folder,
																				   Operation_Scenario_Folder)

						operation_case_folder = app.GetProjectFolder("scen")
						_op_sc_results_folder, _folder_exists2 = create_folder(operation_case_folder,
																			   Operation_Scenario_Folder)

						# Check if folder already exists
						task_auto_name = 'Task_Automation_{}'.format(start1)
						_exists, _task_auto_handle = check_if_object_exists(
							location=study_case_results_folder,
							name=task_auto_name + constants.PowerFactory.autotasks_command)
						if _exists:
							task_automation = _task_auto_handle[0]
						else:
							# Create ComTasks and store in parent_study_case_folder (required location)
							task_automation = create_object(study_case_results_folder, 'ComTasks',
															'Task_Automation_{}'.format(start1))

						cont_count = 0
						# TODO: Swap order so adds filters to base case and then applies contingencies

						# Create results files for this study case so copied as part of copy / paste process
						# Find and create base study case for each project
						base_case_name = '{}_{}'.format(sc_list_parameters[0],
													 constants.HASTInputs.base_case)
						sc_base, op_base = copy_study_case(
							name= base_case_name,
							sc_target_folder=study_case_results_folder,
							sc=study_case,
							op_target_folder = _op_sc_results_folder,
							op=scenario)

						# Create a variations folder for this project so that new mutual impedances created
						# during running can be deleted easily.
						# Create new variation within variations_folder
						variations_folder = app.GetProjectFolder("scheme")
						_variations_folder, _folder_exists3 = create_folder(variations_folder,
																			Variation_Name)
						variation = create_variation(_variations_folder, constants.PowerFactory.pf_scheme,
													 Variation_Name)
						activate_variation(variation)

						# Create and activate recording stage within variation
						_ = create_stage(variation, constants.PowerFactory.pf_stage, Variation_Name)

						# Check base case converges, otherwise skip all contingencies
						cls_base_sc = hast2_1.pf.PFStudyCase(full_name=base_case_name,
															 list_parameters=sc_list_parameters,
															 cont_name=constants.HASTInputs.base_case,
															 sc=sc_base,
															 op=op_base,
															 prj=_prj,
															 task_auto=task_automation,
															 uid=start1,
															 base_case=True,
															 results_pth=Temp_Results_Export)

						# Add load flow command to study case
						cls_base_sc.create_load_flow(load_flow_settings=Load_Flow_Setting)
						# Run load flow and determine if successful fo base case that has been copied as an extra check
						if not cls_base_sc.run_load_flow():
							logger.error(('Load flow not successful for base study case {} and therefore no '
										  'contingencies for this case will be studied')
										 .format(base_case_name))
							# Remove reference to base case
							del cls_base_sc
							# -- POTENTIAL LOOP EXIT COMMAND --
							continue

						# Add study case to dictionary of projects
						if project_name not in prj_dict.keys():
							# Create a new project and add these objects so that they will be deleted once
							# the study has been completed
							# Variations all stored in a folder so that they can be deleted as a group.
							objects_to_delete = [study_case_results_folder,
												 _op_sc_results_folder,
												 _variations_folder]
							prj_dict[project_name] = hast2_1.pf.PFProject(name=project_name,
																		  prj=_prj,
																		  task_auto=task_automation,
																		  folders=objects_to_delete,
																		  include_mutual=Excel_Export_Z12)
						prj_dict[project_name].sc_cases.append(cls_base_sc)
						prj_dict[project_name].sc_base = sc_base

						# Studies running successfully so continue
						create_study_case_results_files(cls_sc=cls_base_sc, cls_prj=prj_dict[project_name])

						# Add filters for base case
						add_all_filters(
							new_filter_list=new_filter_list,
							cont_name=base_case_name,
							sc=sc_base,
							op=op_base,
							sc_target_folder=study_case_results_folder,
							op_target_folder=_op_sc_results_folder,
							variations_folder=_variations_folder,
							list_params=sc_list_parameters,
							cont_short_name=constants.HASTInputs.base_case,
							cls_prj = prj_dict[project_name],
							create_studies=True)

						while cont_count < len(new_contingency_list):
							logger.debug('Carrying out Contingency Check: {}'
										 .format(new_contingency_list[cont_count][0]))
							# sc_base and op_base now reflect base case so no need to check
							if new_contingency_list[cont_count][0] == "Base_Case":
								cont_count += 1
								continue

							cont_name = '{}_{}'.format(sc_list_parameters[0],
													   new_contingency_list[cont_count][0])
							cont_study_case, cont_scenario = copy_study_case(
								name = cont_name,
								sc_target_folder=study_case_results_folder,
								sc=sc_base,
								op_target_folder=_op_sc_results_folder,
								op=op_base
							)

							# Take outages described for contingency
							for _switch in new_contingency_list[cont_count][1:]:
								switch_coup(_switch[0], _switch[1])

							save_active_scenario()

							# Create new class reference with all the details for this contingency and then add to
							# list to be returned
							_study_cls = hast2_1.pf.PFStudyCase(full_name=cont_name,
																list_parameters=sc_list_parameters,
																cont_name=new_contingency_list[cont_count][0],
																sc=cont_study_case,
																op=cont_scenario,
																prj=_prj,
																task_auto=task_automation,
																uid=start1,
																base_case=False,
																results_pth=Temp_Results_Export)
							# Create load flow case and check if error
							_study_cls.create_load_flow(load_flow_settings=Load_Flow_Setting)

							# Only add load flow to study case list and project list if load_flow successful, will still
							# remain in folder of study cases but will be skipped in freq_scan and harmonic lf
							if _study_cls.run_load_flow():

								_study_cls.create_studies(logger=logger,
														  fs=FS_Sim, hldf=HRM_Sim,
														  fs_settings=Fsweep_Settings,
														  hldf_settings=Harmonic_Loadflow_Settings)

								# Add study case to file
								prj_dict[project_name].sc_cases.append(_study_cls)
							else:
								logger.error(('Load flow for study case {} not successful and so frequency scans ' +
											 ' and harmonic load flows will not be run for this case')
											 .format(cont_study_case))

							# Add filter for every contingency
							add_all_filters(
								new_filter_list=new_filter_list,
								cont_name=cont_name,
								sc=cont_study_case,
								op=cont_scenario,
								sc_target_folder=study_case_results_folder,
								op_target_folder=_op_sc_results_folder,
								variations_folder=_variations_folder,
								list_params=sc_list_parameters,
								cont_short_name=new_contingency_list[cont_count][0],
								cls_prj=prj_dict[project_name],
								create_studies=True)

							# TODO: Use enumerator rather than iterating counter
							cont_count = cont_count + 1

						cls_base_sc.sc.Activate()
						cls_base_sc.create_studies(logger=logger,
												   fs=FS_Sim, hldf=HRM_Sim,
												   fs_settings=Fsweep_Settings,
												   hldf_settings=Harmonic_Loadflow_Settings)

					else:
						logger.error('Problem with initial load flow: {}'.format(sc_list_parameters[0]))
				else:
					logger.error('Problem with Scenario: {} {}'.format(sc_list_parameters[0],
																	   sc_list_parameters[3]))
			else:
				logger.error('Problem with Studycase: {} {}'
							 .format(sc_list_parameters[0], sc_list_parameters[2]))
		else:
			logger.error('Problem Activating Project: {} {}'
						 .format(sc_list_parameters[0], sc_list_parameters[1]))
	msg1 = '-- Finished Checking Study Cases in {:.2f}'.format(time.clock() - time_sc_check)
	msg2 = '\t'+'_'*100+'\t'
	logger.info('{}\n{}'.format(msg1,msg2))
	return prj_dict

def check_terminals(list_of_points, prj_name, sc_name): 		# This checks and creates the list of terminals to add to the Results file
	"""
		Creates list of terminals to be added to the results file
	:param list list_of_points:  List of terminal references defined in excel_writing.TerminalDetails
	:param str prj_name:  Name of current project terminals being looked for in
	:param str sc_name:  Name of study case currently being looked for in
	:return: (list, bool) (terminals_index, terminals_error): (list of terminal indexes, 
															1 if error adding any of the terminals)
	"""
	# Constants used to determine successfully running
	# terminals_error is set to 1 if fails
	terminals_error = 0
	# terminals_index contains list of the terminal handles to be used in the study
	terminals_index = list()
	for terminal in list_of_points:
		# Get handle for the substation that contains the data
		pf_sub = app.GetCalcRelevantObjects(terminal.substation)

		if len(pf_sub) == 0:
			# If the length is 0 then it means no items are returned
			logger.error(('Substation entry {} does not exist or is not active in the PowerFactory Project: {}'
						  ' and Study Case: {}.\n'
						  'Therefore not possible to obtain results for User named terminal: {}')
						 .format(terminal.substation, prj_name, sc_name, terminal.name))
			terminals_ok = 1
			# Continue with next terminal
			continue

		# Check if terminal is contained within substation
		# Get list of all terminals that match this name
		terminals_in_substation = pf_sub[0].GetContents(terminal.terminal)

		# Confirm that at least 1 terminal with the required named exists in the substation

		if len(terminals_in_substation) == 0:
			logger.error(('Terminal entry {} does not exist or is not active in the Substation {} for the PowerFactory '
						  'Project: {} and Study Case: {}\n'
						  'Therefore not possible to obtain results for HAST User Terminal Name: {}')
						 .format(terminal.terminal, terminal.substation, prj_name, sc_name, terminal.terminal))
			terminals_ok = 1
			continue

		# If multiple terminals with the same name exist then alert User.  This should not be possible in the current
		# version of PowerFactory
		if len(terminals_in_substation) > 1:
			logger.warning(('More than 1 terminal with the name {} found in substation {} for PowerFactory Project {} '
						   'and Study Case {} associated with the User Terminal Input {}.\n'
						   'Results will only be returned for the 1st one called {}')
						   .format(terminal.terminal, terminal.substation, prj_name, sc_name, terminal.name,
								   terminals_in_substation[0]))

		# Add PowerFactory handle reference to terminal class <excel_writing.TerminalDetails>
		terminal.pf_handle = terminals_in_substation[0]
		# Added to list of terminals_index for backwards compatibility
		# TODO: Rather than adding to list here just add handle reference to class and deal with when needed
		terminals_index.append([terminal.name, terminal.substation, terminal.terminal,
								terminal.pf_handle, terminal.include_mutual])

	# All terminals have been added so print list of terminals considered and warn user if number of terminals
	# considered is not the same as the number intended
	if len(list_of_points) != len(terminals_index):
		logger.warning(('The HAST Inputs spreadsheet requested results from {} terminals but only {} terminals have '
					   'been found in PowerFactory').format(len(list_of_points), len(terminals_index)))

	logger.info('The following Terminals are used for the analysis:')
	for item in terminals_index:
			logger.info('\t - HAST Name: {}, Substation: {}, Terminal: {}, PF Reference: {}'
						.format(item[0], item[1], item[2], item[3]))

	# Returns list of terminals and a status flag = 1 if there was an error finding any of the terminals
	return terminals_index, terminals_error

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
		# TODO:  Skip for other options such as 0 in Substation name, or
		if item[0] == "Base_Case":														# Skips the base case
			list_of_couplers.append("Base_Case")
			list_of_couplers.append(0)
		else:
			list_of_couplers.append(item[0])
			for aa in item[1:]:
				coupler_exists = False
				# TODO: Adjust to ensure that if aa[0] is an integer it will still work i.e. '{}.ElmSubstat'.format(aa[0])
				# TODO:  This will need careful checking for contingency error in case is the Base Case entry.
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
	msg1 = 'Contingencies Used for Analysis:'
	msg2 = '\n'.join(['\t - Name: {}'.format(x[0]) for x in new_contingency_list])
	logger.info('{}\n{}'.format(msg1, msg2))
	return new_contingency_list, contingencies_ok

def check_filters(list_of_filters):			# Checks and creates list of terminals to add the Filters to
	"""
			Checks the status of each contingency
		:param list list_of_filters: List of filters to be checked where each filter is of type excel_writing.FilterDetails
		:return: List of filters to actually be studied 
		"""
	filters_ok = True
	for item in list_of_filters:  # This loops through the contingencies to find the couplers
		if not item.include:
			continue

		skip_filter = False
		filter_name = item.name
		substation = item.sub
		terminal = item.term
		hdl_substation = app.GetCalcRelevantObjects(substation)		# Finds the Substation

		#If substation exists then find relevant terminal in substation contents
		if len(hdl_substation) == 0:									# If it doesn't find it it reports it and skips it
			logger.warning(
				'For filter: {}, Python substation entry for {} does not exist in case so not modelled'
					.format(filter_name, substation))
			item.include = False
		else:
			# Find terminal in substation if doesn't exist then raise error and skip to next item
			hdl_terminal = hdl_substation[0].GetContents(terminal)
			if len(hdl_terminal) == 0:
				logger.warning(
					'For filter: {}, Python terminal {} within substation {} does not exist in the case and so not modelled'
						.format(filter_name, substation, terminal))
				item.include = False
				continue

		# Get nominal voltage of terminal as nominal voltage for filter
		item.nom_voltage = hdl_terminal[0].GetAttribute(constants.PowerFactory.pf_term_voltage)


	new_filters_list = [_x for _x in list_of_filters if _x.include]
	if len(new_filters_list) == 0:
		logger.info('No filters to include')
		filters_ok = False

	return new_filters_list, filters_ok

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

def main(import_workbook, results_export_folder=None, uid=None, include_nom_voltage=True):
	"""
		Ensures this code is only run when run as the main script and not for unittesting
	:param str import_workbook: Target HAST workbook for import
	:param str results_export_folder: (optional=None) String to results export folder,
		if None provided then uses the value in the HAST inputs or the existing folder
	:param str uid: (optional=None) String to use for unique identified,
		if None provided then based on running time of script
	:param bool include_nom_voltage:  (optional=True) - If set to False then nominal voltage is removed from
		list of variables
	:return:
	"""
	# Removes nominal voltage from list of variables to ensure backwards compatibility
	if not include_nom_voltage:
		try:
			constants.HASTInputs.fs_term_variables.remove(constants.PowerFactory.pf_nom_voltage)
		except ValueError:
			# Must have already been removed from list
			pass
		Process_HAST_extract.INCLUDE_NOM_VOLTAGE = False
	# In case it is removed and needs adding back in again
	elif constants.PowerFactory.pf_nom_voltage not in constants.HASTInputs.fs_term_variables:
		constants.HASTInputs.fs_term_variables.append(constants.PowerFactory.pf_nom_voltage)


	sys.path.append(DIG_PATH)
	sys.path.append(DIG_PYTHON_PATH)

	os.environ['PATH'] = os.environ['PATH'] + ';' + DIG_PATH
	title = ('::::::::::::::::::::::::::::::::::::::::::::::::::::::::::\n' +
		'NAME:           Harmonics Automated Simulation Tool (HAST)\n' +
		'VERSION:        {}\n' +
		'AUTHOR:         {}, {}, {}\n' +
		'STATUS:			{}\n' +
		'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::\n')\
		.format(__name__, __author__, __email__, __phone__, __status__)

	import powerfactory  # Power factory module imported here to allow running in unattended mode

	# Start Timer
	start = time.clock()
	global start1
	if uid is None:
		start1 = (time.strftime("%y_%m_%d_%H_%M_%S"))
	else:
		start1 = uid

	# Power factory commands
	# --------------------------------------------------------------------------------------------------------------------
	# TODO:  Need to add in capability here to capture script fail and release so that powerfactory license is released
	global app

	# TODO: Write unittest to check exception raised if powerfactory not loaded
	if distutils.version.StrictVersion(powerfactory.__version__) > distutils.version.StrictVersion('17.0.0'):
		# powerfactory after 2017 has an error handler when trying to load
		try:
			app = powerfactory.GetApplicationExt()  # Start PowerFactory  in engine  mode
		except powerfactory.ExitError as error:
			print('An error occured trying to start PowerFactory, there error was {}'.format(error))
			print('Error Code returned by PowerFactory = {}'.format(error.code))
			raise SyntaxError('Power Factory Load Error - Unable to run HAST')
	else:
		# In case of an older version of PowerFactory being run
		app = powerfactory.GetApplication()
		if app is None:
			print('Unable to load PowerFactory')
			raise SyntaxError('PowerFactory Load Error - Unable to run HAST')

	# Get commands from PowerFactory that are used in multiple locations
	app.ClearOutputWindow()  # Clear Output Window

	global Variation_Name
	Variation_Name = "Temporary_Variation" + start1

	# Create excel instance to deal with retrieving import data from excel
	# TODO: Make use of class in <hast2.excel_writing> for complete processing of inputs
	with hast2.excel_writing.Excel(print_info=print1, print_error=print2) as excel_cls:
		# Reads in the Settings from the spreadsheet
		analysis_dict = excel_cls.import_excel_harmonic_inputs(workbookname=import_workbook)
	# TODO: Complete processing to convert everything to use class for processing
	cls_hast_inputs = hast2.excel_writing.HASTInputs(hast_inputs=analysis_dict, uid_time=start1)

	study_settings = analysis_dict[constants.HASTInputs.study_settings]
	if len(study_settings) != 20:
		print2('Error, Check input Study Settings there should be 20 Items in the list there are only: {} {}'
			   .format(len(study_settings), study_settings))
	if results_export_folder is None:
		if not study_settings[0]:
			# If there is no output location in the spreadsheet it sets it to the CWD
			results_export_folder = filelocation
		else:
			# Folder to Export Excel Results too
			results_export_folder = study_settings[0]

	# Temporary results folder to store in as exported (create if doesn't exist)
	global Temp_Results_Export
	Temp_Results_Export = os.path.join(results_export_folder, start1)
	if not os.path.exists(Temp_Results_Export):
		# Try/except statement added in case poor path given in HAST Inputs that cannot be found or created
		try:
			os.makedirs(Temp_Results_Export)
		except (FileNotFoundError, OSError):
			# Since folder cannot be created assume issue with full path provided
			temp_results_export_original = Temp_Results_Export
			results_export_folder = filelocation
			Temp_Results_Export = os.path.join(results_export_folder, start1)
			# Have to use print statement since logger has not been enabled yet
			print(('Not able to create the following folder for saving the PowerFactory raw results:  {}\n'
				   'Instead the raw results will be saved in: {}\n'
				   'and the processed results will be saved in: {}')
				  .format(temp_results_export_original, Temp_Results_Export, results_export_folder))
			# Additional Try / Except clause in case fails on next attempt as well such as because read / write
			# permissions not possible
			try:
				os.makedirs(Temp_Results_Export)
			except (FileNotFoundError, OSError):
				print(('Unable to create a folder in {} either and so the script will stop!\n'
					   'Check the HAST Inputs file: {}')
					  .format(Temp_Results_Export, import_workbook))
				raise FileNotFoundError('Unable to create folder for PowerFactory Results')


	# Declare file names
	excel_results = os.path.join(results_export_folder, study_settings[1] + start1)			# Name of Exported Results File
	progress_log = os.path.join(results_export_folder, study_settings[2] + start1 + ".txt")	# Progress File
	error_log = os.path.join(results_export_folder, study_settings[3] + start1 + ".txt")	# Error File
	debug_log = os.path.join(results_export_folder + 'DEBUG' + start1 + '.txt')

	# Setup logger with reference to powerfactory app
	global logger
	logger = hast2.logger.Logger(pth_debug_log=debug_log,
								 pth_progress_log=progress_log,
								 pth_error_log=error_log,
								 app=app,
								 debug=DEBUG_MODE)

	# Disable graphic updating
	if not DEBUG_MODE and logger.pf_executed:
		logger.info('Graphic updating and detailed load flow results will not be shown until script completes')
		app.SetGraphicUpdate(0)
		app.EchoOff()
	else:
		logger.info('Running in debug mode and so all details and progress updates are output')

	shutil.copy(src=import_workbook, dst=os.path.join(Temp_Results_Export, 'HAST Inputs_{}.xlsx'.format(start1)))

	# Random_Log = results_export_folder + "Random_Log_" + start1 + ".txt"		# For printing random info solely for development
	global Net_Elm
	Net_Elm = study_settings[4]													# Where all the Network elements are stored
	global Mut_Elm_Fld
	Mut_Elm_Fld = study_settings[5]												# Name of the folder to create under the network elem to store mutual impedances
	global Operation_Scenario_Folder
	Operation_Scenario_Folder = study_settings[7]	+ start1 					# Name of the folder to store Operational Scenarios
	global Pre_Case_Check
	Pre_Case_Check = study_settings[8]											# Checks the N-1 for load flow convergence and saves operational scenarios.
	global FS_Sim
	FS_Sim = study_settings[9]													# Carries out Frequency Sweep Analysis
	global HRM_Sim
	HRM_Sim = study_settings[10]												# Carries out Harmonic Load Flow Analysis
	global Skip_Unsolved_Ldf
	Skip_Unsolved_Ldf = study_settings[11]										# Skips the frequency sweep if the load flow doesn't solve
	global Delete_Created_Folders
	Delete_Created_Folders = study_settings[12]									# Deletes the Results folder, Mutual Elements and the Operational Scenario folder
	export_to_excel = study_settings[13]										# Export the results to Excel
	global Excel_Export_Z12
	Excel_Export_Z12 = study_settings[18]										# Export Mutual Impedances to excel


	c = constants.PowerFactory
	logger.info(title)
	for keys,values in analysis_dict.items():									# Prints all the inputs to progress log
		logger.info(keys)
		logger.info(values)
	global List_of_Studycases
	List_of_Studycases = analysis_dict[c.sht_Scenarios]						# Uses the list of Studycases
	if len(List_of_Studycases) <1:												# Check there are the right number of inputs
		logger.error('Error - Check excel input Base_Scenarios there should be at least 1 Item in the list')
	global List_of_Contingencies
	List_of_Contingencies = analysis_dict[c.sht_Contingencies]						# Uses the list of Contingencies
	if len(List_of_Contingencies) <1:											# Check there are the right number of inputs
		logger.error('Error - Check excel input Contingencies there should be at least 1 Item in the list')
	global List_of_Filters
	List_of_Filters = cls_hast_inputs.list_of_filters
	if len(List_of_Filters) == 0:
		logger.info('No harmonic filters listed for studies')
	global List_of_Points
	# Uses the class with list of terminals to allow referencing by name rather than position
	List_of_Points = cls_hast_inputs.list_of_terms
	if len(List_of_Points) <1:													# Check there are the right number of inputs
		logger.error('Error - Check excel input Terminals there should be at least 1 Item in the list')
	global Load_Flow_Setting
	Load_Flow_Setting = analysis_dict[c.sht_LF]						# Imports Settings for LDF calculation
	if len(Load_Flow_Setting) != 55:											# Check there are the right number of inputs
		print2('Error - Check excel input Loadflow_Settings there should be 55 Items in the list there are only: {} {}'
			   .format(len(Load_Flow_Setting), Load_Flow_Setting))
	global Fsweep_Settings
	Fsweep_Settings = analysis_dict[c.sht_Freq]							# Imports Settings for Frequency Sweep calculation
	if len(Fsweep_Settings) != 16:												# Check there are the right number of inputs
		print2('Error - Check excel input Frequency_Sweep there should be 16 Items in the list there are only: {} {}'
			   .format(len(Fsweep_Settings), Fsweep_Settings))
	global Harmonic_Loadflow_Settings
	Harmonic_Loadflow_Settings = analysis_dict[c.sht_HLF]				# Imports Settings for Harmonic LDF calculation
	if len(Harmonic_Loadflow_Settings) != 15:									# Check there are the right number of inputs
		logger.error(('Error - Check excel input Harmonic_Loadflow Settings there should be 17 Items in the list '
					 'there are only: {} {}')
					 .format(len(Harmonic_Loadflow_Settings), Harmonic_Loadflow_Settings))

	# Check all study cases converge, etc. and produce a new study case + operational scenario for each one
	# Adjusted to now return a list of handles to class <hast22.pf.PF_Study_Case> which contain handles for the powerfactory
	# scenario objects that require activating.
	dict_of_projects = check_list_of_studycases(List_of_Studycases)
	if len(dict_of_projects) == 0:
		logger.critical('No base cases converged and so no studies to run.  Check your inputs and that you have a '
						'convergent model.\n'
						'See above for a list of non-convergent load flows')
		logger.flush()
		logger.close_logging()
		del logger
		raise RuntimeError('No projects successfully initialised for harmonic studies to be run')



	# Confirms that studies should be run and that there are some projects that have actually been successfully produced
	# If all base load flows are non-convergent then there are no studies to run.
	if FS_Sim or HRM_Sim:
		# Get and deactivate current project
		current_prj = app.GetActiveProject()
		current_prj.Deactivate()

		t1 = time.clock()
		# As of v2.1.2 tested on multiple projects and seems to be producing the output correctly
		for prj_name, prj_cls in dict_of_projects.items():
			# Activate project
			prj_activation_failed = prj_cls.prj.Activate()

			#If failed activation then returns 1 (i.e. True) and for loop is continued
			if prj_activation_failed == 1:
				logger.error('Not possible to activate project {} and so no studies are performed for this project'
							 .format(prj_cls.name))
				continue

			t1_prj_start = time.clock()
			logger.info('Creating studies for Study Cases associated with project {}'.format(prj_cls.name))
			prj_cls.update_auto_exec(fs=FS_Sim, hldf=HRM_Sim)

			logger.info('Creating of commands for studies in project {} completed in {:0.2f} seconds'
						.format(prj_cls.name, time.clock()-t1_prj_start))
			t1_prj_studies = time.clock()

			logger.info('Parallel running of frequency scans and harmonic load flows associated with project {}'
						.format(prj_cls.name))

			# Task Auto Execute seems to break logger so flush progress and error
			# log commands here and then retrieve again
			logger.flush()

			# Call Task automation to run studies
			# TODO:  Sometimes seem to hit an error where license does not allow this to run.  Need to check why that
			# TODO: is occurring and figure out how to avoid.  Potential would be to close / open PF
			prj_cls.task_auto.Execute()

			logger.info('Studies for project {} completed in {:0.2f} seconds'
						.format(prj_cls.name, time.clock()-t1_prj_studies))

			# Once studies complete, deactivate project
			prj_cls.prj.Deactivate()

		logger.info('PowerFactory studies all completed in {:0.2f} seconds'.format(time.clock()-t1))

	# Deleting newly created folders which will include study_cases and operational_scenarios
	if Delete_Created_Folders:
		t_start_delete = time.clock()
		logger.info('Deleting newly created folders as part of this study')
		for project, prj_cls in dict_of_projects.items():
			logger.debug('Deleting items for project: {}'.format(prj_cls.name))
			# Activate project
			prj_cls.prj.Activate()
			# Deactivate currently active study case so that items from project can be deleted
			deactivate_study_case()
			# Loop through each folder and try to delete
			for folder in prj_cls.folders:
				delete_object(folder)

			# TODO: After deleting folders will be useful to reactivate original study case, operating scenario and
			# TODO: variations.
			prj_cls.prj.Deactivate()
		logger.info('Deletion of created items completed in {:.2f} seconds'.format(time.clock() - t_start_delete))

	# Plot results to excel
	if export_to_excel:
		combined_df, vars_in_hast = Process_HAST_extract.combine_multiple_hast_runs(
			search_pths=[Temp_Results_Export],
			drop_duplicates=False
		)
		Process_HAST_extract.extract_results(
			pth_file=excel_results + constants.ResultsExtract.extension,
			df=combined_df,
			vars_to_export=vars_in_hast)

	logger.info('Total Time: {:.2f}'.format(time.clock() - start))

	# Close the logger since script has now completed and this forces flushing of the open logs before script exits
	logger.flush()
	logger.close_logging()

	del logger
	app = None

	return excel_results + constants.ResultsExtract.extension

if __name__ == '__main__':
	# Determine whether to use GUI for file selection or just HAST_Inputs.xlsx file
	Import_Workbook = filelocation + hast_inputs_filename  # Gets the CWD current working directory

	# Determine if a file already exists in the script folder that conforms to the standard input of
	# HAST_Inputs.xlsx, if not then ask user to select the file.
	if not os.path.isfile(Import_Workbook):
		if tk:
			# Load GUI for user to select file for HAST imports
			Import_Workbook = hast2.gui.file_selector(
				initial_pth=filelocation,
				open_file=True,
				lbl_file_select='Select HAST Inputs.xlsx file',
				openfile_types=(('XLSX files','*.xlsx'),
								('All Files','*.*'))
			)
			if Import_Workbook == '':
				raise IOError('User did not select a file for HAST Inputs.xlsx')
			else:
				# gui.file_selector returns a list of which first element is the file required
				Import_Workbook = Import_Workbook[0]
		else:
			# Captures an error with both the file not existing and the GUI not loading
			raise IOError('No HAST_Inputs.xlsx and GUI unable to run. Place HAST_Inputs.xlsx in script directory')


	# Run HAST with the selected Import Workbook
	Results_File = main(Import_Workbook)