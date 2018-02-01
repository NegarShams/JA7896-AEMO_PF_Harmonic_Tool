Title = ("""::::::::::::::::::::::::::::::::::::::::::::::::::::::::::\n
NAME:             HAST Harmonics Automated Simulation Tool\n
VERSION:          1.2 [24 April 2017]\n
AUTHOR:           Barry O'Connell\n
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::\n""")

# IMPORT SOME PYTHON MODULES
## ---------------------------------------------------------------------------------------------------------------------------
import os,sys
DIG_PATH = """C:\\Program Files\\DIgSILENT\\PowerFactory 2016 SP3\\"""
sys.path.append(DIG_PATH)
os.environ['PATH'] = os.environ['PATH'] + ';' +  DIG_PATH 

# Important Notes
## ---------------------------------------------------------------------------------------------------------------------------
# Notepad ++ is a useful tool for viewing python coded
# Install Python 3.5 to your C:\ or D:\ drive. Do not install in C:\\programfiles  as the win32com module needs write access to create a cache and it wont have that in program files
#Use these commands to check Environment variables are setup correctly
#help("modules")						# Check if powerfactory is in your modules, if not copy powerfactory python dll (c:\\programfiles\\pf etc) to python directory (eg C:\\python3.5\\DLL) 
#print(os.environ["PATH"])				# Check to see the correct path above was appended to your environment variables
#for param in os.environ.keys():
#   print "%20s %s" % (param,os.environ[param])

# If you are having trouble with numpy scipy ensure that you either install the modules or anaconda which has these modules present
# You can comment out numpy and scipy in the import section if you set Excel_Convex_Hull = False. This will then skip creating the points for the convex hull

import powerfactory 					# Power factory module see notes above
import time                          	# Time
import ctypes                        	# For creating startup message box
import win32com.client              	# Windows COM clients needed for excel etc. if having trouble see notes
import math								# 
import numpy as np						# install anaconda it has numpy in it  https://www.continuum.io/downloads
from scipy.spatial import ConvexHull	# install anaconda it has scipy in it  https://www.continuum.io/downloads
import re								# Used for stripping text strings
#import shutil
#import inspect                      # Inspect functions
#import string                       # Processing text
#import operator
#import textwrap

# Start Timer
filelocation = os.getcwd() + "\\"
start = time.clock()
start1 = (time.strftime("%y_%m_%d_%H_%M_%S")) 

# Excel commands
xl = win32com.client.gencache.EnsureDispatch('Excel.Application')   # Call dispatch excel application excel

# Power factory commands
#--------------------------------------------------------------------------------------------------------------------------------
app = powerfactory.GetApplication() 							# Start PowerFactory  in engine  mode
#help("powerfactory")
user = app.GetCurrentUser()										# Get the current active user
ldf = app.GetFromStudyCase("ComLdf")							# Get load flow command
hldf = app.GetFromStudyCase("ComHldf")							# Get Harmonic load flow
frq = app.GetFromStudyCase("ComFsweep")							# Get Frequency Sweep Command
ini = app.GetFromStudyCase("ComInc")							# Get Dynamic Initialisation
sim = app.GetFromStudyCase("ComSim")							# Get Dynamic Simulation
shc = app.GetFromStudyCase("ComShc")							# Get short circuit command
res = app.GetFromStudyCase("ComRes")							# Get Result Export Command
wr = app.GetFromStudyCase("ComWr")								# Get Write command for wmf and bmp files
app.ClearOutputWindow()											# Clear Output Window

# Functions --------------------------------------------------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------------------------------------------------------------	
def print1(bf, name, af):			# Used to print a message to both python, PF and write it to the file with double space
	name = str(name)
	print(name)
	#app.PrintError(str message)	# Prints message as an error
	#app.PrintInfo(str message)		# Prints message as info
	#app.PrintWarn(str message)		# Prints message as a warning
	app.PrintPlain(name)	# Prints message as plain
	Progress = open(Progress_Log, "a")		# Progress File
	Progress.write(bf*"\n")	
	Progress.write(name)
	Progress.write(af*"\n")	
	Progress.close()
	return

def print2(name):			# Used to print error message to both python, PF and write it to the file 
	global Error_Count
	bf = 2
	af = 0
	name = str(name)
	print(name)
	app.PrintError(name)	# Prints message as an error
	Progress = open(Progress_Log, "a")		# Progress File
	Progress.write(bf*"\n")	
	Progress.write("Error No." + str(Error_Count) + " " + name)
	Progress.write(af*"\n")	
	Progress.close()
	Error = open(Error_Log, "a")
	Error.write(bf*"\n")
	Error.write("Error No." + str(Error_Count) + " " + name)
	Error.write(af*"\n")
	Error.close()
	Error_Count = Error_Count + 1
	return

def print3(bf, name, af):		# Used to print a message to both python, PF and write it to the file with double space
	name = str(name)
	print(name)
	#app.PrintError(str message)		# Prints message as an error
	#app.PrintInfo(str message)			# Prints message as info
	#app.PrintWarn(str message)			# Prints message as a warning
	app.PrintPlain(name)				# Prints message as plain
	Random = open(Random_Log, "a")		# Progress File
	Random.write(bf*"\n")	
	Random.write(name)
	Random.write(af*"\n")	
	Random.close()
	return
	
def startup_message():		# Used to Create a startup dialog box
    app.PrintPlain("\ndef startup_message\n")
    reply = ctypes.windll.user32.MessageBoxA(0, "Close all PSSe and Excel Files \nThis program kills these tasks while running \
                                                \nOK to proceed", "Attention", 1)
    if reply == 2:
        sys.exit("Cancel Run")
    return

def Import_Excel_Harmonic_Inputs(workbookname):		# Import Excel Harmonic Input Settings
	xl = win32com.client.gencache.EnsureDispatch('Excel.Application')   # Call disptach excel application excel
	wb = xl.Workbooks.Open(workbookname)                                # Open workbook
	c = win32com.client.constants                                       #
	xl.Visible = False                                                  # Make excel Visible
	xl.DisplayAlerts = False                                            # Don't Display Alerts
	Analysis_Sheets = (("Study_Settings", "B5"), ("Base_Scenarios", "A5"), ("Contingencies","A5"), ("Terminals","A5"), 
					("Loadflow_Settings","D5"), ("Frequency_Sweep","D5"), ("Harmonic_Loadflow","D5"))
	analysis_dict = {}

	for x in Analysis_Sheets:
		ws = wb.Sheets(x[0])                                                # Set Active Sheet
		sh = ws.Activate()                                                  # Activate Sheet
		cell_start = x[1]                                                   # Starting Cell
		
		ws.Range(cell_start).End(c.xlDown).Select()                         # Equivalent to shift end down
		row_end = xl.Selection.Address
		row_input = []
		if x[0] == "Contingencies" or x[0] == "Base_Scenarios" or x[0] == "Terminals":	# For these sheets
			cell_start_alph = re.sub('[^a-zA-Z]', '', cell_start)								# Gets the starting cell alpha C5 = C
			cell_start_num = int(re.sub('[^\d\.]', '', cell_start))								# Gets the starting cell number C5 = 5
			cell_end = int(re.sub('[^\d\.]', '', row_end))										# Gets the ending cell alpha E5 = E
			cell_range_num  = list(range(cell_start_num,(cell_end+1)))							# Gets the ending cell number E5 = 5
			check_value = ws.Range(cell_start_alph + str(cell_start_num + 1)).Value				# Check the value below cell called
			if check_value == None:																# If the cell is None 
				aces = [cell_start]	
			else:
				aces = [cell_start_alph + str(no) for no in cell_range_num]							# 
			count_row = 0
			while count_row < len(aces):
				ws.Range(aces[count_row]).End(c.xlToRight).Select()
				col_end = xl.Selection.Address                                      # Returns address of last cells
				row_value = ws.Range(aces[count_row] + ":" + col_end).Value
				row_value = row_value[0]
				if x[0] == "Contingencies":
					if len(row_value) > 2:
						aa = row_value[1:]
						station_name = aa[0::3]
						breaker_name = aa[1::3]
						breaker_status = aa[2::3]
						breaker_name1 = [str(nam) + ".ElmCoup" for nam in breaker_name]
						aa1 = list(zip(station_name, breaker_name1,breaker_status))
						aa1.insert(0,row_value[0])
					else:
						aa1= [row_value[0],[0]]
					row_value = aa1
				if x[0] == "Base_Scenarios":
					row_value = [row_value[0], row_value[1], row_value[2] + ".IntCase", row_value[3] + ".IntScenario"]
				if x[0] == "Terminals":
					row_value = [row_value[0], row_value[1] + ".ElmSubstat", row_value[2] + ".ElmTerm"]
				row_input.append(row_value)
				count_row = count_row + 1
		elif x[0] == "Study_Settings" or x[0] == "Loadflow_Settings" or x[0] == "Frequency_Sweep" or x[0] == "Harmonic_Loadflow":
			row_value = ws.Range(cell_start + ":" + row_end).Value
			for item in row_value:
				row_input.append(item[0])
			if x[0] == "Loadflow_Settings":
				z = row_input
				row_input = [int(z[0]), int(z[1]), int(z[2]), int(z[3]), int(z[4]), int(z[5]), int(z[6]), int(z[7]), int(z[8]), float(z[9]),
							int(z[10]), int(z[11]), int(z[12]), z[13], z[14],
							int(z[15]), int(z[16]), int(z[17]), int(z[18]), float(z[19]), int(z[20]), float(z[21]),
							int(z[22]), int(z[23]), int(z[24]), int(z[25]), int(z[26]), int(z[27]), int(z[28]),
							z[29], z[30], int(z[31]), z[32], int(z[33]),
							int(z[34]), int(z[35]), int(z[36]), int(z[37]), z[38], z[39], z[40], z[41], int(z[42]), z[43],
							z[44], z[45], z[46], z[47], z[48], z[49], z[50], z[51], int(z[52]), int(z[53]), int(z[54])]
			if x[0] == "Frequency_Sweep":
				z = row_input
				row_input = [z[0], z[1], int(z[2]), z[3], z[4], z[5], int(z[6]), z[7], z[8], z[9],
							z[10], z[11], z[12], z[13], int(z[14]), int(z[15])]
			if x[0] == "Harmonic_Loadflow":
				z = row_input
				row_input = [int(z[0]), int(z[1]), int(z[2]), int(z[3]), z[4], z[5], z[6], z[7],
							z[8], int(z[9]), int(z[10]), int(z[11]), int(z[12]), int(z[13]), int(z[14])]
		analysis_dict[(x[0])] = row_input      # Imports range of values into a list of lists
	
	wb.Close()                                                          # Close Workbook           
	return analysis_dict

def ActivateProject(Project): 		# Activate project
	pro = app.ActivateProject(Project) 										# Activate project
	if pro == 0:															# Project Activate Successfully
		print1(1, "Activated Project Successfully: " + str(Project), 0)		# Print Information to progress log and PowerFactory window
		prj = app.GetActiveProject()										# Get active project
	else:																	# Project Failed to Activate
		print2(("Error Unsuccessfully Activated Project: " + str(Project) + "................................")) 	# Print Information to progress log and PowerFactory window and Error Log
		prj = []
	return prj

def ActivateStudyCase(StudyCase): 		# Activate Study case
	DeactivateStudyCase()
	StudyCaseFolder1 = app.GetProjectFolder("study")							# Returns string the location of the project folder for study cases, scen, 
	StudyCase1 = StudyCaseFolder1.GetContents(StudyCase)
	if len(StudyCase1) > 0:
		cas = StudyCase1[0].Activate() 														# Activate Study case
		if cas == 0:
			print1(1, "Activated Study Case Successfully: " + str(StudyCase1[0]), 0)			
		else:
			print2(("Error Unsuccessfully Activated Study Case: " + str(StudyCase) + " ................................"))
	else:
		print2("Couldn't Activate Studycase as none matching name in case: " + str(StudyCase))
		cas = 1
		StudyCase1 = [[]]
	return StudyCase1[0], cas

def DeactivateStudyCase(): 		# Deactivate Scenario
	Study = app.GetActiveStudyCase()
	if Study is not None:
		sce = Study.Deactivate() 											# Deactivate Study case
		if sce == 0:
			pass
			#print1(1,"Deactivated Active Study Case Successfully : " + str(Study),0)
		elif sce > 0:
			print2(("Error Unsuccessfully Deactivated Study Case: " + str(Study) + " ................................"))
			print2(("Unsuccessfully Deactivated Scenario Error Code: " + str(sce)))
	elif Study is None:
		print1(2,"No Study Case Active to Deactivate ................................",0)
	return 
	
def ActivateScenario(Scenario): 		# Activate Scenario
	ScenarioFolder1 = app.GetProjectFolder("scen")							# Returns string the location of the project folder for study cases, scen, 
	Scenario1 = ScenarioFolder1.GetContents(Scenario)
	DeactivateScenario()
	sce = Scenario1[0].Activate() 											# Activate Study case
	if sce == 0:
		print1(1,"Activated Scenario Successfully: " + str(Scenario1[0]),0)
	elif sce > 0:
		print2(("Error Unsuccessfully Activated Scenario: " + str(Scenario) + " ................................"))
		print2(("Unsuccessfully Activated Scenario Error Code: " + str(sce)))
	return Scenario1[0], sce

def ActivateScenario1(Scenario): 		# Activate Scenario
	sce = Scenario.Activate() 											# Activate Study case
	if sce == 0:
		print1(1,"Activated Scenario Successfully: " + str(Scenario),0)
	elif sce == 1:
		print2(("Error Unsuccessfully Activated Scenario: " + str(Scenario) + " ................................"))
		print2(("Unsuccessfully Activated Scenario Error Code: " + str(sce)))
	return sce
	
def DeactivateScenario(): 		# Deactivate Scenario
	Scenario1 = app.GetActiveScenario()
	if Scenario1 is not None:
		sce = Scenario1.Deactivate() 											# Deactivate Study case
		if sce == 0:
			pass
			#print1(1,("Deactivated Active Scenario Successfully : " + str(Scenario1)),0)
		elif sce > 0:
			print2(("Error Unsuccessfully Deactivated Scenario: " + str(Scenario1) + " ................................"))
			print2(("Unsuccessfully Deactivated Scenario Error Code: " + str(sce)))
	else:
		print1(2,"No Scenario Active to Deactivate ................................",0)
	return 

def SaveActiveScenario():		# Saves the current active Operational Scenario
	Scenario1 = app.GetActiveScenario()
	sce = Scenario1.Save()
	if sce == 0:
		print1(1,("Saved Active Scenario Successfully: " + str(Scenario1)),0)
	elif sce == 1 and Scenario1 is None:
		print2(("Error Unsuccessfully Saved Scenario: " + str(Scenario1) + " ................................"))
		print2(("Unsuccessfully Saved Scenario Error Code: " + str(sce)))
	elif Scenario1 is None:
		print1(2,"No Scenario Active to Save ................................",0)
	return 
	
def GetActiveVariations():			# Get Active Network Variations
	Variations =  app.GetActiveNetworkVariations()
	print1(2,"Current Active Variations: ",0)
	if len(Variations) > 1:
		for item in Variations:
			aa = str(item)
			pp = aa.split("Variations.IntPrjfolder\\")
			ss = pp[1]
			tt = ss.split(".IntScheme")
			print1(1,tt[0],0)
	elif len(Variations) == 1:
		print1(1,Variations,0)
	else:
		print1(1,"No Variations Active",0)
	return Variations

def Create_Variation(Folder,pfclass,name):
	Variation = Create_Object(Folder,pfclass,name)
	Variation.icolor = 1
	#Variation.chr_name = "1"
	#Variation.for_name = "1"
	#Variation.desc = ["1","2"]
	#Variation.dat_src = "1"
	app.PrintPlain(Variation)
	return Variation

def ActivateVariation(Variation): 		# Activate Scenario
	sce = Variation.Activate() 											# Activate Study case
	if sce == 0:
		print1(1,"Activated Variation Successfully: " + str(Variation),0)
	elif sce == 1:
		print2(("Error Unsuccessfully Activated Variation: " + str(Variation) + " ................................"))
		print2(("Unsuccessfully Activated Variation Error Code: " + str(sce)))
	return sce

def Create_Stage(location,pfclass,name):
	stage = location.CreateObject(pfclass,name)
	stage.loc_name = name
	#stage.cDate = 1/1/2014
	#stage.cTime = 
	#stage.chr_name = "1"
	#stage.for_name =
	#stage.desc = 
	#stage.dat_src = 
	#stage.appr_status = 
	#stage.InvCosts = 
	#stage.AddCosts = 
	#stage.OrigVal = 
	#stage.SrcVal = 
	#stage.LifeSpan = 
	Activate_Stage(stage)
	return stage

def Activate_Stage(stage):
	sce = stage.Activate()
	if sce == 0:
		print1(1,"Activated Variation Stage Successfully: " + str(stage),0)
	elif sce != 0:
		print2(("Error Unsuccessfully Activated Variation Stage: " + str(stage) + " ................................"))
		print2(("Unsuccessfully Activated Variation Stage Error Code: " + str(sce)))
	return
	
def LoadFlow(load_flow_settings):		# Inputs load flow settings and executes load flow
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
	#ldf.rembar = load_flow_settings[13] # Reference Busbar
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
	#ldf.outcmd =  load_flow_settings[41]          		# Output
	ldf.iopt_chctr = load_flow_settings[42]    			# Check Control Conditions
	#ldf.chkcmd = load_flow_settings[43]            	# Command

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
		print1(1,"Load Flow calculation successful, time taken: " + str(round(t2,2)) + " seconds",0)
	elif error_code == 1:
		print2("Load Flow failed due to divergence of inner loops, time taken: " + str(round(t2,2)) + " seconds................................")
	elif error_code == 2:
		print2("Load Flow failed due to divergence of outer loops, time taken: " + str(round(t2,2)) + " seconds................................")
	return error_code

def HarmLoadFlow(results,Harmonic_Loadflow_Settings):		# Inputs load flow settings and executes load flow
	t1 = time.clock()
	## Loadflow settings
	## -------------------------------------------------------------------------------------
	# Basic
	hldf.iopt_net = Harmonic_Loadflow_Settings[0]               	# Calculation method (0 Balanced AC, 1 Unbalanced AC, DC)
	hldf.iopt_allfrq = Harmonic_Loadflow_Settings[1]				# Calculate Harmonic Load Flow 0 - Single Frequency 1 - All Frequencies
	hldf.iopt_flicker = Harmonic_Loadflow_Settings[2] 				# Calculate Flicker
	hldf.iopt_SkV = Harmonic_Loadflow_Settings[3] 					# Calculate Sk at Fundamental Frequency		
	hldf.frnom = Harmonic_Loadflow_Settings[4]            			# Nominal Frequency
	hldf.fshow = Harmonic_Loadflow_Settings[5]             			# Output Frequency
	hldf.ifshow = Harmonic_Loadflow_Settings[6]  					# Harmonic Order
	hldf.p_resvar = results          								# Results Variable
	#hldf.cbutldf =  Harmonic_Loadflow_Settings[8]               	# Load flow
	
	# IEC 61000-3-6
	hldf.iopt_harmsrc = Harmonic_Loadflow_Settings[9]				# Treatment of Harmonic Sources
	
	# Advanced Options
	hldf.iopt_thd = Harmonic_Loadflow_Settings[10] 					# Calculate HD and THD 0 Based on Fundamental Frequency values 1 Based on rated voltage/current
	hldf.maxHrmOrder = Harmonic_Loadflow_Settings[11] 				# Max Harmonic order for calculation of THD and THF
	hldf.iopt_HF = Harmonic_Loadflow_Settings[12] 					# Calculate Harmonic Factor (HF)
	hldf.ioutall = Harmonic_Loadflow_Settings[13] 					# Calculate R, X at output frequency for all nodes
	hldf.expQ = Harmonic_Loadflow_Settings[14] 						# Calculation of Factor-K (BS 7821) for Transformers
	
	error_code = hldf.Execute()
	t2 = time.clock() - t1
	if error_code == 0:
		print1(1,"Harmonic Load Flow calculation successful: " + str(round(t2,2)) + " seconds",0)
	elif error_code > 0:
		print2("Harmonic Load Flow calculation unsuccessful: " + str(round(t2,2)) + " seconds................................")
	return error_code
	
def FSweep(results,fsweep_settings):		# Inputs Frequency Sweep Settings and executes sweep
	t1 = time.clock()
	## Frequency Sweep Settings
	## -------------------------------------------------------------------------------------
	# Basic
	nomfreq = fsweep_settings[0]                  # Nominal Frequency
	maxfrq = fsweep_settings[1]                 	# Maximum Frequency
	frq.iopt_net = fsweep_settings[2]                # Network Representation (0=Balanced 1=Unbalanced)
	frq.fstart = fsweep_settings[3]                	# Impedance Calculation Start frequency
	frq.fstop = fsweep_settings[4]              # Stop Frequency
	frq.fstep = fsweep_settings[5]                 # Step Size
	frq.i_adapt = fsweep_settings[6]                 # Automatic Step Size Adaption
	frq.frnom = fsweep_settings[7]             # Nominal Frequency
	frq.fshow = fsweep_settings[8]              # Output Frequency
	frq.ifshow = fsweep_settings[9]   # Harmonic Order
	frq.p_resvar = results          # Results Variable
	#frq.cbutldf = fsweep_settings[11]                 # Load flow

	# Advanced
	frq.errmax = fsweep_settings[12]               # Setting for Step Size Adaption    Maximum Prediction Error
	frq.errinc = fsweep_settings[13]              # Minimum Prediction Error
	frq.ninc = fsweep_settings[14]                   # Step Size Increase Delay
	frq.ioutall = fsweep_settings[15]                 # Calculate R, X at output frequency for all nodes

	error_code = frq.Execute()	
	t2 = time.clock() - t1
	if error_code == 0:
		print1(1,"Frequency Sweep calculation successful, time taken: " + str(round(t2,2)) + " seconds",0)
	elif error_code > 0:
		print2("Frequency Sweep calculation unsuccessful, time taken: " + str(round(t2,2)) + " seconds................................")
	return error_code

def SaveOpScenario(name,active): 	# Saves an operational Scenario
	scenario = app.SaveAsScenario(name, active)	# name of scenario and 1 to activate it after or 0 to not activate
	if len(str(scenario)) == 0:
		print2("Scenario : " + str(name) + " save unsuccessful" + " ..............................................")	
	else:
		print1(2,"Scenario : " + str(name) + " saved successfully",0)
	return scenario

def Switch_Coup(element,service):			# Switches an Coupler out if 0 in if 1
	element.on_off = service
	if service == 0:
		print1(1,"Switching Element: Out of service " + str(element),0)
	if service == 1:
		print1(1,"Switching Element: In to service " + str(element),0)
	return	

def Check_If_Folder_Exists(location, name):		# Checks if the folder exists
	new_object = location.GetContents(name + ".IntFolder")
	folder_exists = 0
	if len(new_object) > 0:
		print1(2,"Folder already exists: " + str(name),0)
		folder_exists = 1
	return new_object, folder_exists

def Create_Folder(location, name):		# Creates Folder in location
	print1(1,"Creating Folder: " + str(name),0)	
	folder1, folder_exists = Check_If_Folder_Exists(location, name)				# Check if the Results folder exists if it doesn't create it using date and time
	if folder_exists == 0:
		new_object = location.CreateObject("IntFolder",name)
		loc_name = name							# Name of Folder
		owner = "Barry"							# Owner
		iopt_sys = 0							# Attributes System
		iopt_type = 0							# Folder Type 0 Common
		for_name = ""							# Foreign key
		desc = ""								# Description			
	else:
		new_object = folder1[0]
	return new_object, folder_exists

def Delete_Folder(location, name):		# Deletes Folder in Location
	new_object = location.GetContents(name + ".IntFolder",)
	if len(new_object) > 0:
		new_object[0].Delete()
	return	

def Create_Mutual_Impedance_List(location,Terminal_List):		# Creates a mutual Impedance list from the terminal list in a folder under the active studycase
	print1(1,"Creating: Mutual Impedance List of Terminals",0)
	Terminal_List1 = list(Terminal_List)
	List_of_Mutual = []
	count = 0
	for y in Terminal_List1:
		for x in Terminal_List1:
			if x[3] != y[3]:
				elmmut = Create_Mutual_Elm(location,(str(y[0]) + "_" + str(x[0])),y[3],x[3])
				List_of_Mutual.append([str(y[0]),(str(y[0]) + "_" + str(x[0])),elmmut,y[3],x[3]])
	return List_of_Mutual
	
def Create_Mutual_Elm(location,name,bus1,bus2):		# Creates Mutual Impedance between two terminals
	#elmmut = app.GetFromStudyCase(name + )				# Get relevant object or create if it doesn't exist
	elmmut = Create_Object(location, "ElmMut", name)
	elmmut.loc_name = name
	elmmut.bus1 = bus1
	elmmut.bus2 = bus2
	elmmut.outserv = 0
	return elmmut

def Get_Object(object):			# retrieves an object based on filter strings
	ob1 = app.GetCalcRelevantObjects(object)
	return ob1

def Delete_Object(object):			# retrieves an object based on filter strings
	ob1 = object.Delete()
	if ob1 == 0:
		print1(1,("Object Successfully Deleted: " + str(object)),0)
	else:
		print2(("Error Deleting Object: " + str(object) + "................................"))
	return

def Check_If_Object_Exists(location, name):  	# Check if the object exists
	print1(2,[location, name],0)
	new_object = location.GetContents(name)
	object_exists = 0
	if len(new_object) > 0:
		print1(2,"Object Exists: " + str(name),0)
		object_exists = 1
	return object_exists, new_object	
	
def Add_Copy(folder,object,name1):		# copies an object to a new folder Name 1 = new name 
	new_object = folder.AddCopy(object, name1)
	if new_object is not None:
		print1(1,("AddCopy Successful: " + str(object)),0)
	else:
		print2(("Error AddCopy Unsuccessful: " + str(object) + " to " + str(newfolder) + " as name " + name1))
	return new_object

def Create_Object(location,pfclass,name):			# Creates a database object in a specified loaction of a specified class
	new_object = location.CreateObject(pfclass,name)	
	return new_object

def Create_Results_File(location, name, type):			# Creates Results File
	# Manipulating Results Files
	#sweep = app.GetFromStudyCase(name)				# Get relevant object or create if it doesn't exist (Old way more explicit now)
	sweep = Create_Object(location, "ElmRes", name)
	#sweep.Delete()									# Deletes results object
	p = sweep.Clear()								# Clears Data
	variable_contents = sweep.GetContents()			# Gets the existing variables
	#app.PrintPlain(variable_contents)				# Prints the existing variables
	for cc in variable_contents:					# Loops through and deletes the existing variables
		cc.Delete()
	sweep.calTp = type								# Frequency / Harmonic
	sweep.header = ["Hello Barry"]
	sweep.desc = ["Barry Description"]
	return sweep

def Check_List_of_Studycases(list):		# Check List of Projects, Study Cases, Operational Scenarios, 
	print1(2,"__________________________________________________________________________________________________________________________________",0)
	print1(2,"Checking all Projects, Study Cases and Scenarios Solve for Load Flow, it will also check N-1 and create the operational scenarios if Pre_Case_Check is True\n",0)
	count_studycase = 0
	new_list =[]
	err = 0
	while count_studycase < len(list):
		prj = ActivateProject(List_of_Studycases[count_studycase][1])												# Activate Project
		if len(str(prj)) > 0:
			StudyCase, study_error = ActivateStudyCase(list[count_studycase][2])									# Activate Case
			if study_error == 0:
				Scenario, scen_err = ActivateScenario(list[count_studycase][3])										# Activate Scenario
				if scen_err == 0:
					ldf_err = LoadFlow(Load_Flow_Setting)																			# Perform Load Flow
					if ldf_err == 0 or Skip_Unsolved_Ldf == False:
						new_list.append(list[count_studycase])
						print1(2,"Studycase Scenario Solving added to analysis list: " + str(list[count_studycase]),0)
						if Pre_Case_Check == True:																	# Checks all the contingencies and terminals are in the prj,cas
							New_Contingency_List, Con_ok = Check_Contingencis(List_of_Contingencies) 				# Checks to see all the elements in the contingency list are in the case file
							Terminals_index, Term_ok = Check_Terminals(List_of_Points)								# Checks to see if all the terminals are in the case file skips any that aren't
							Operation_Case_Folder = app.GetProjectFolder("scen")	
							op_sc_results_folder, folder_exists2 = Create_Folder(Operation_Case_Folder, Operation_Scenario_Folder)
							cont_count = 0
							while cont_count < len(New_Contingency_List):
								print1(2,"Carrying out Contingency Pre Stage Check: " + New_Contingency_List[cont_count][0],0)
								DeactivateScenario()																# Can't copy activated Scenario so deactivate it
								new_scenario = Add_Copy(op_sc_results_folder,Scenario,List_of_Studycases[count_studycase][0] + str("_" + New_Contingency_List[cont_count][0]))	# Copies the base scenario	
								scen_error = ActivateScenario1(new_scenario)										# Activates the base scenario
								if New_Contingency_List[cont_count][0] != "Base_Case":								# Apply Contingencies if it is not the base case
									for switch in New_Contingency_List[cont_count][1:]:																								
										Switch_Coup(switch[0],switch[1])
								SaveActiveScenario()
								ldf_err_cde = LoadFlow(Load_Flow_Setting)															# Carry out load flow
								cont_count = cont_count + 1
					else:
						print2("Problem with Loadflow: " + str(list[count_studycase][0]))
				else:
					print2("Problem with Scenario: " + str(list[count_studycase][0]) + " " + str(list[count_studycase][3]))
			else:
				print2("Problem with Studycase: " + str(list[count_studycase][0]) + " " + str(list[count_studycase][2]))
		else:
			print2("Problem Activating Project: " + str(list[count_studycase][0]) + " " + str(list[count_studycase][1]))		
		count_studycase = count_studycase + 1
	print1(1,"Finished Checking Study Cases",0)
	print1(2,"__________________________________________________________________________________________________________________________________",2)
	return new_list

def Check_Terminals(List_of_Points): 		# This checks and creates the list of terminals to add to the Results file
	All_Terminals_ok = 0
	Terminals_index = []														# Where the calculated variables will be stored
	tm_count = 0
	while tm_count < len(List_of_Points):										# This loops through the variables adding them to the results file													
		t = app.GetCalcRelevantObjects(List_of_Points[tm_count][1])				# Finds the Substation 
		if len(t) == 0:															# If it doesn't find it it reports it and skips it
			print2("Python substation entry for does not exist in case: " + List_of_Points[tm_count][1] + "..............................................")
		else:
			t1 = t[0].GetContents()													# Gets the Contents of the substations (ie objects) 
			terminal_exists = False
			for t2 in t1:															# Gets the contents of the objects in the Substaion
				if List_of_Points[tm_count][2]  in str(t2):												# Checks to see if the terminal is there
					Terminals_index.append([List_of_Points[tm_count][0],List_of_Points[tm_count][1],List_of_Points[tm_count][2],t2])					# Appends Terminals ( Name, Terminal Name, Terminal object data)
					terminal_exists = True											# Marks that it found the terminal
			if terminal_exists == False:											# Flags to the user the terminal didn't exist
				print2("Python Entry does not exist in case: " + List_of_Points[tm_count][2] + ".ElmTerm ..............................................")
				All_Terminals_ok = 1
		tm_count = tm_count + 1
	print1(2,"Terminals Used for Analysis: ",0)
	tm_count = 0
	while tm_count < len(List_of_Points):
		print1(1,List_of_Points[tm_count],0)
		tm_count = tm_count + 1
	return Terminals_index, All_Terminals_ok	
	
def Check_Contingencis(List_of_Contingencies): 		# This checks and creates the list of terminals to add to the Results file
	All_Terminals_ok = 0
	New_Contingency_List = []															# Where the calculated variables will be stored
	for item in List_of_Contingencies:													# This loops through the contingencies to find the couplers				
		skip_contingency = False
		list_of_couplers = []
		if item[0] == "Base_Case":														# Skips the base case
			coupler_exists = True
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
								elif aa[2] == "Close":
									breaker_operation = 1
								else:
									print2("Contingency entry: " + item[0] + ". Coupler in Substation: " + aa[0] +  " " + aa[1] + " could not carry out: " + aa[2] + " ..............................................")
								list_of_couplers.append([t2,breaker_operation])									# Appends Terminals ( Name, Terminal Name, Terminal object data)
						coupler_exists = True										# Marks that it found the terminal		
					if coupler_exists == False:
						print2("Contingency entry: " + item[0] + ". Coupler does not exist in Substation: " + aa[0] +  " " + aa[1] + " ..............................................")
						print2("Skipping Contingency")
						skip_contingency = True
		if skip_contingency == True:											# Flags to the user the terminal didn't exist
			All_Terminals_ok = 1
		elif skip_contingency == False:
			New_Contingency_List.append(list_of_couplers)
	print1(2,"Contingencies Used for Analysis:",0)
	for item in New_Contingency_List:
		print1(1,item,0)
	return New_Contingency_List, All_Terminals_ok
	
def Add_Vars_Res(elmres,element,Vars):	# Adds the results variables to the results file
	if len(Vars) > 1:
		for x in Vars:
			elmres.AddVariable(element,x)
	elif len(Vars) == 1:
		elmres.AddVariable(element,Vars[0])
	return

def Plot(name, type, results, terminal, variable, description, clear):	# Plots the results in Powerfactory
	setDesktop=app.GetGraphicsBoard()
	viPage = setDesktop.GetPage((name+"_plt"),1) 				# Searches and activates an existing plot, if it does not exist then it will overwrite it 0 or 1
	oVi = viPage.GetVI(name+"_plt",type,1)						# Name of Virtual Instrument panel, type, create if it doesn't exist
	if clear == 0:	
		oVi.Clear()												# Clears the existing visplot
	oVi.AddResVars(results, terminal, variable)					# Adds Results File, Element and Variables to the plot
	oVi.SetCrvDesc((clear+1),description)
	viPage.DoAutoScaleX()
	viPage.DoAutoScaleY()
	return

def Results_Export(results,output):		# Not used Export results file into Excel
	res.pResult = results					# Export from
	res.iopt_exp = 6 						# Type of File
	res.f_name = output						# File Name
	res.iopt_sep = 1 						# Use System Separators
	res.iopt_newx = 0						# Number of time points not needed if you choose export all variables
	res.iopt_honly = 0						# Export Object header only
	res.iopt_csel = 0 						# Variable Selection
	#res.resultobj
	res.Execute()
	
def Retrieve_Results(elmres,type):			# Reads results into python lists from results file
	# Note both column number and row start at 1.
	# The first column is usually the scale ie timestep, frequency etc.
	# The columns are made up of Objects from left to right (ElmTerm, ElmLne)
	# The Objects then have sub variables (m:R, m:X etc)
	elmres.Load()
	cno = elmres.GetNumberOfColumns()	# Returns number of Columns 
	rno = elmres.GetNumberOfRows()		# Returns number of Rows in File
	Results = []
	for i in range(cno):
		column = []
		p = elmres.GetObject(i) 		# Object
		d = elmres.GetVariable(i)		# Variable
		column.append(d)
		column.append(str(p))
		#column.append(d)
		# app.PrintPlain([i,p,d])	
		for j in range(rno):
			r,t = elmres.GetValue(j,i)
			#app.PrintPlain([i,p,d,j,t])
			column.append(t)
		Results.append(column)
	if type == 1:
		Results = Results[:-1]
	Scale = Results[-1:]
	Results = Results[:-1]
	elmres.Release()
	return Scale[0], Results

def ReadTextfile(file):		# Reads in Textfile
	text_file = open(file, "r")
	content = text_file.readlines()
	text_file.close()
	print1(2,"Reading in textfile: " + str(textfile),0)
	return content
		
def Create_Workbook(workbookname):			# Create Workbook
	print1(2,"Creating Workbook: " + workbookname,0)
	xl = win32com.client.gencache.EnsureDispatch('Excel.Application')   # Call dispatch excel application excel
	c = win32com.client.constants                                       # used for retrieving constants from excel
	wb = xl.Workbooks.Add()                                             # Add workbook
	xl.Visible = Excel_Visible                                          # Make excel Visible
	xl.DisplayAlerts = False                                            # Don't Display Alerts
	#wb.Sheets(1).Delete()                                              # Delete Sheet 1 ie "Sheet 1"
	#ws = wb.Worksheets.Add()                                           # Add worksheet
	wb.SaveAs(workbookname)                                             # Save Workbook
	return wb
	
def Create_Sheet_Plot(Sheet_Name, FS_Results, HRM_Results, wb):      # Extract information from out file
	t1 = time.clock()
	Sheet_Name = Sheet_Name
	print1(1,"Creating Sheet: " + Sheet_Name,0)
	xl = win32com.client.gencache.EnsureDispatch('Excel.Application')   # Call disptach excel application excel
	c = win32com.client.constants                                       #
	ws = wb.Worksheets.Add()                                            # Add worksheet      
	startrow = 2
	startcol = 1
	newrow = 2
	newcol = 1
	
	R_First, R_Last, R_First, R_Last, X_First, X_Last, Z_First, Z_Last, Z_12_First, Z_12_Last, HRM_endrow, HRM_First, HRM_Last = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
	if Excel_Export_RX == True:
		startcol = 19
	if Excel_Export_Z == True or Excel_Export_HRM == True:
		startrow = 33
	if Excel_Export_Z == True and Excel_Export_HRM == True:
		startrow = 62
	
	if FS_Sim == True:
		if Excel_Export_RX == True or Excel_Export_Z == True or Excel_Export_Z12 == True:		# Prints the FS Scale 
			endrow = startrow + len(FS_Results[0]) - 1
			endcol = startcol + len(FS_Results) - 1
			# Plots the Scale_________________________________________________________________________________________________________
			scale = FS_Results[0]
			scale_end = scale[-1]
			ws.Range(ws.Cells(startrow,startcol),ws.Cells(endrow,startcol)).Value = list(zip(*[FS_Results[0]]))
			newcol = startcol + 1
			FS_Results = FS_Results[1:] # Remove scale
		
		if Excel_Export_RX == True:		# Export the RX data and graphs the Impedance Loci
			# Insert R data in excel______________________________________________________________________________________________
			newcol = newcol
			R_First = newcol
			R_Results, X_Results = [], []
			for x in FS_Results:													 
				if x[0] == "m:R":
					ws.Range(ws.Cells(startrow,newcol),ws.Cells(endrow,newcol)).Value = list(zip(*[x])) 
					R_Results.append(x[3:]) 
					newcol = newcol + 1	
			R_Last = newcol - 1
			
			# Insert X data in excel______________________________________________________________________________________________
			newcol = newcol + 1
			X_First = newcol
			for x in FS_Results:													 
				if x[0] == "m:X":
					ws.Range(ws.Cells(startrow,newcol),ws.Cells(endrow,newcol)).Value = list(zip(*[x])) 
					X_Results.append(x[3:]) 
					newcol = newcol + 1	
			X_Last = newcol	- 1	
			
			t2 = time.clock() - t1
			print1(1,"Inserting RX data self impedance data, time taken: " + str(round(t2,2)) + " seconds",0)		
			t1 = time.clock()
			
			# Graph R X Data impedance Loci_______________________________________________________________________________________
			chart_width = 400		# Width of Graph
			chart_height = 300		# Height of Chart
			left = 30
			top = 45				# Top Starting Point
			if Excel_Export_Z == True or Excel_Export_HRM == True:
				top = startrow * 15
			graph_across = 1		# Number of Graphs Across the Page
			graph_spacing = 25		# Spacing between the graphs
			noofgraphrows = math.ceil(len(scale[3:])/graph_across) - 1
			noofgraphrowsrange = list(range(0,noofgraphrows+1))
			gph_coord = []									# List of Graph coordinates for Impedance Loci
			for uyt in noofgraphrowsrange:					# This creates the graph coordinates
				mnb = list(range(0,graph_across))
				for lkj in mnb:
					gph_coord.append([(left + lkj*(chart_width + graph_spacing)),(top + uyt*(chart_height + graph_spacing))])
			
			# This section is used to calculate the position of the rows for non Integer Harmonics
			scale_list_int = []
			scale_clipped = scale[3:] 			# Remove Headers
			lp_count = 0
			for lkp in scale_clipped:			# Get position of harmonics
				hjg = (lkp/50).is_integer()
				if hjg == True:
					scale_list_int.append(lp_count)
					#app.PrintPlain(lp_count)
				lp_count = lp_count + 1
			if len(scale_list_int) < 3:
				print2("The frequency range you have given is less than 2 integer harmonics") 
			else:
				diff = (scale_list_int[2] - scale_list_int[1]) / 2	# Get the difference between positions of whole harmonics
			
			Non_int_rows = []
			for wer in scale_list_int:			# Plot the 1st point of range and the position of the actual harmonic and the end of harmonic range [75, 100, 125] would return [0,1,2]
				if diff < 1:
					Pr = wer
					Qr = wer
				else:
					Pr = int(wer - diff)
					Qr = int(wer + diff)
				if Pr < 0:
					Pr = 0
				if Qr > (len(scale_clipped) - 1):
					Qr = len(scale_clipped) - 1
				Non_int_rows.append([Pr,wer,Qr])
	
			gc = 0 					# Graph Count
			new_row = startrow + 3
			X_Results = list(zip(*X_Results))
			R_Results = list(zip(*R_Results))
			startrow1 = (endrow + 3)
			startcol1 = startcol + 3
			for hrm in Non_int_rows:											# Plots the Graphs for the Harmonics including non integer rows
				ws.Range(ws.Cells(1,1),ws.Cells(1,2)).Select()					# Important for computation as it doesn't graph all the selection first ie these cells should be blank
				ch = ws.Shapes.AddChart(c.xlXYScatter,gph_coord[gc][0],gph_coord[gc][1],chart_width,chart_height).Select()      # AddChart(Type, Left, Top, Width, Height)
				xl.ActiveChart.ApplyLayout(1)																					# Select Layout 1-11
				xl.ActiveChart.ChartTitle.Characters.Text = " Harmonic Order " + str(int(scale_clipped[hrm[1]]/50))      # Add Title
				xl.ActiveChart.SeriesCollection(1).Delete()
				#xl.ActiveChart.Legend.Delete()                                                         # Delete legend
				xl.ActiveChart.Axes(c.xlCategory).AxisTitle.Text = "Resistance in Ohms"                 # X Axis
				xl.ActiveChart.Axes(c.xlCategory).MinimumScale = 0                                      # Set minimum of x axis
				xl.ActiveChart.Axes(c.xlCategory).TickLabels.NumberFormat = "0"                         # Set number of decimals 0.0                                               
				xl.ActiveChart.Axes(c.xlValue).AxisTitle.Text = "Reactance in Ohms"                     # Y Axis
				xl.ActiveChart.Axes(c.xlValue).TickLabels.NumberFormat = "0"                            # Set number of decimals        
				impedance_loci_pos = list(range(R_First,R_Last,2))	
				RX_Con = []
				
				# This is used to graph non integer harmonics on the same plot as integer
				for tres in range(hrm[0],(hrm[2]+1)):			
					series = xl.ActiveChart.SeriesCollection().NewSeries()
					series.XValues = ws.Range(ws.Cells((startrow + 3 + tres) ,R_First),ws.Cells((startrow + 3 + tres),R_Last))		# X Value
					series.Values = ws.Range(ws.Cells((startrow + 3 + tres),X_First),ws.Cells((startrow + 3 + tres),X_Last))		# Y Value
					series.Name = ws.Cells((startrow + 3 + tres),startcol)															# Name
					series.MarkerSize = 5																							# Marker Size
					series.MarkerStyle = 3																							# Marker type
					prop_count = 0
					if tres < len(R_Results):
						for desd in R_Results[tres]:
							RX_Con.append([desd,X_Results[tres][prop_count]])
							prop_count = prop_count + 1
					else:
						print2("The Scale is longer then the dataset it probably means that you have selected automatic step size adaption in FSweep")
				
				if Excel_Convex_Hull == True:										# This is used to the convex hull of the points on the graph with a line
					RXarray = np.array(RX_Con)										# Converts the RX data to a numpy array 
					convexRX = ConvexHull1(RXarray)									# Get the min area points of the array needs to be in numpy array
					endcol1 = (startcol1+ len(convexRX[0]) - 1)	
					ws.Range(ws.Cells(startrow1,startcol1),ws.Cells(startrow1,endcol1)).Value = convexRX[0]				# Adds R data to Excel
					ws.Range(ws.Cells(startrow1+1,startcol1),ws.Cells(startrow1+1,endcol1)).Value = convexRX[1]			# Add X data to Excel
					series = xl.ActiveChart.SeriesCollection().NewSeries()												# Adds a new series for it
					series.XValues = ws.Range(ws.Cells(startrow1,startcol1),ws.Cells(startrow1,endcol1))				# X Value
					series.Values = ws.Range(ws.Cells(startrow1+1,startcol1),ws.Cells(startrow1+1,endcol1))				# Y Value												# Name	
					ws.Cells(startrow1,startcol).Value = str(int(scale_clipped[hrm[1]])) + " Hz"
					ws.Cells(startrow1,startcol + 1).Value = str(int(scale_clipped[hrm[0]])) + " Hz"
					ws.Cells(startrow1+1,startcol + 1).Value = str(int(scale_clipped[hrm[2]])) + " Hz"
					ws.Cells(startrow1,startcol + 2).Value = "R"
					ws.Cells(startrow1+1,startcol + 2).Value = "X"			
					series.MarkerSize = 5																				# Marker Size
					series.MarkerStyle = -4142																			# Marker type
					series.Format.Line.Visible = True																	# Marker line
					series.Format.Line.ForeColor.RGB = 12611584															# Colour is red + green*256 + blue*256*256
					series.Name = "Convex Hull"																			# Name		
					
					# Plots the graphs for the customers
					ws.Range(ws.Cells(1,1),ws.Cells(1,2)).Select()					
					ch = ws.Shapes.AddChart(c.xlXYScatter,gph_coord[gc][0] + 425,gph_coord[gc][1],chart_width,chart_height).Select()      # AddChart(Type, Left, Top, Width, Height)
					xl.ActiveChart.ApplyLayout(1)																					# Select Layout 1-11
					xl.ActiveChart.ChartTitle.Characters.Text = " Harmonic Order " + str(int(scale_clipped[hrm[1]]/50))      # Add Title
					xl.ActiveChart.SeriesCollection(1).Delete()
					xl.ActiveChart.Axes(c.xlCategory).AxisTitle.Text = "Resistance in Ohms"                 # X Axis
					xl.ActiveChart.Axes(c.xlCategory).MinimumScale = 0                                      # Set minimum of x axis
					xl.ActiveChart.Axes(c.xlCategory).TickLabels.NumberFormat = "0"                         # Set number of decimals 0.0                                               
					xl.ActiveChart.Axes(c.xlValue).AxisTitle.Text = "Reactance in Ohms"                     # Y Axis
					xl.ActiveChart.Axes(c.xlValue).TickLabels.NumberFormat = "0"                            # Set number of decimals       
					series = xl.ActiveChart.SeriesCollection().NewSeries()												# Adds a new series for it
					series.XValues = ws.Range(ws.Cells(startrow1,startcol1),ws.Cells(startrow1,endcol1))				# X Value
					series.Values = ws.Range(ws.Cells(startrow1+1,startcol1),ws.Cells(startrow1+1,endcol1))				# Y Value												# Name
					series.Name = ws.Cells(startrow1,startcol)															# Name
					series.MarkerSize = 5																				# Marker Size
					series.MarkerStyle = -4142																			# Marker type
					series.Format.Line.Visible = True																	# Marker line
					series.Format.Line.ForeColor.RGB = 12611584															# Colour is red + green*256 + blue*256*256					

				startrow1 = startrow1 + 2
				new_row = new_row + 1
				gc = gc + 1		
			t2 = time.clock() - t1
			print1(1,"Graphing RX data self impedance data, time taken: " + str(round(t2,2)) + " seconds",0)		
			t1 = time.clock()
			newcol = newcol + 1
		ws.Name = Sheet_Name                # Rename worksheet          
		wb.Save()							# Save Workbook		
		if Excel_Export_Z == True:		# Export Z data and graphs
			# Insert Z data in excel_______________________________________________________________________________________________
			ws.Range(ws.Cells(startrow,newcol),ws.Cells(endrow,newcol)).Value = list(zip(*[scale]))
			if Excel_Export_RX == True:
				newcol = newcol + 1
			Z_First = newcol - 1
			Z_Results, Base_Case_Pos = [], []
			for x in FS_Results:													 
				if x[0] == "m:Z":
					ws.Range(ws.Cells(startrow,newcol),ws.Cells(endrow,newcol)).Value = list(zip(*[x])) 
					Z_Results.append(x[3:]) 
					if x[2] == "Base_Case":
						Base_Case_Pos.append([newcol])
					newcol = newcol + 1	
			Z_Last = newcol - 1	
			t2 = time.clock() - t1
			print1(1,"Inserting Z self impedance data, time taken: " + str(round(t2,2)) + " seconds",0)
			t1 = time.clock()
			
			# Graph Z Data_________________________________________________________________________________________________________
			
			if len(Base_Case_Pos) > 1:			# If there is more than 1 Base Case then plot all the bases on one graph and then each base against its N-1 across else just plot them all on one graph.	
				z_no_of_contingencies = int(Base_Case_Pos[1][0]) - int(Base_Case_Pos[0][0])
				ws.Range(ws.Cells(1,1),ws.Cells(1,2)).Select()			# Important for computation as it doesn't graph all the selection first ie these cells should be blank
				ch = ws.Shapes.AddChart(c.xlXYScatterLinesNoMarkers,30,45,825,400).Select()    	# AddChart(Type, Left, Top, Width, Height)
				xl.ActiveChart.ApplyLayout(1)													# Select Layout 1-11
				xl.ActiveChart.ChartTitle.Characters.Text = Sheet_Name + " Base Cases m:Z Self Impedances"	# Add Title
				#xl.ActiveChart.Legend.Delete()                                                	# Delete legend
				xl.ActiveChart.Axes(c.xlCategory).AxisTitle.Text = "Frequency in Hz"            # X Axis
				xl.ActiveChart.Axes(c.xlCategory).MinimumScale = 0                            	# Set minimum of x axis
				xl.ActiveChart.Axes(c.xlCategory).MaximumScale = scale_end                  	# Set maximum of x axis
				xl.ActiveChart.Axes(c.xlCategory).TickLabels.NumberFormat = "0"               	# Set number of decimals 0.0                                               
				xl.ActiveChart.Axes(c.xlValue).AxisTitle.Text = "Impedance in Ohms"      		# Y Axis
				xl.ActiveChart.Axes(c.xlValue).TickLabels.NumberFormat = "0"                   	# Set number of decimals    
				xl.ActiveChart.SeriesCollection(1).Delete()
				
				for zb_col in Base_Case_Pos:			
					series_name1 = ws.Range(ws.Cells((startrow + 1), zb_col[0]),ws.Cells((startrow + 2), zb_col[0])).Value
					series_name = str(series_name1[0][0]) + "_" + str(series_name1[1][0])
					series = xl.ActiveChart.SeriesCollection().NewSeries()
					series.Values = ws.Range(ws.Cells((startrow + 3), zb_col[0]),ws.Cells((endrow), zb_col[0]))						# Y Value
					series.XValues = ws.Range(ws.Cells((startrow + 3), Z_First),ws.Cells((endrow), Z_First))
					series.Name = series_name
				
				zb_count = 1
				for zb_col1 in Base_Case_Pos:	
					ws.Range(ws.Cells(1,1),ws.Cells(1,2)).Select()			# Important for computation as it doesn't graph all the selection first ie these cells should be blank
					ch = ws.Shapes.AddChart(c.xlXYScatterLinesNoMarkers, 30 + zb_count * 855,45,825,400).Select()    	# AddChart(Type, Left, Top, Width, Height)
					xl.ActiveChart.ApplyLayout(1)													# Select Layout 1-11
					series_name1 = ws.Range(ws.Cells((startrow + 1), zb_col1[0]),ws.Cells((startrow + 2), zb_col1[0])).Value
					series_name = str(series_name1[0][0])
					xl.ActiveChart.ChartTitle.Characters.Text = Sheet_Name + " " + str(series_name) + " m:Z Self Impedances"	# Add Title
					#xl.ActiveChart.Legend.Delete()                                                	# Delete legend
					xl.ActiveChart.Axes(c.xlCategory).AxisTitle.Text = "Frequency in Hz"            # X Axis
					xl.ActiveChart.Axes(c.xlCategory).MinimumScale = 0                            	# Set minimum of x axis
					xl.ActiveChart.Axes(c.xlCategory).MaximumScale = scale_end                  	# Set maximum of x axis
					xl.ActiveChart.Axes(c.xlCategory).TickLabels.NumberFormat = "0"               	# Set number of decimals 0.0                                               
					xl.ActiveChart.Axes(c.xlValue).AxisTitle.Text = "Impedance in Ohms"      		# Y Axis
					xl.ActiveChart.Axes(c.xlValue).TickLabels.NumberFormat = "0"                   	# Set number of decimals    
					xl.ActiveChart.SeriesCollection(1).Delete()
					for zzcol in list(range(zb_col1[0],(zb_col1[0] + z_no_of_contingencies))):			
						series_name1 = ws.Range(ws.Cells((startrow + 1), zzcol),ws.Cells((startrow + 2), zzcol)).Value
						series_name = str(series_name1[0][0]) + "_" + str(series_name1[1][0])
						series = xl.ActiveChart.SeriesCollection().NewSeries()
						series.Values = ws.Range(ws.Cells((startrow + 3), zzcol),ws.Cells((endrow), zzcol))						# Y Value
						series.XValues = ws.Range(ws.Cells((startrow + 3), Z_First),ws.Cells((endrow), Z_First))
						series.Name = series_name
					zb_count = zb_count + 1
			
			else:
				ws.Range(ws.Cells(startrow+1,Z_First),ws.Cells(endrow,Z_Last)).Select()			# Important for computation as it doesn't graph all the selection first ie these cells should be blank
				ch = ws.Shapes.AddChart(c.xlXYScatterLinesNoMarkers,30,45,825,400).Select()    	# AddChart(Type, Left, Top, Width, Height)
				xl.ActiveChart.ApplyLayout(1)													# Select Layout 1-11
				xl.ActiveChart.ChartTitle.Characters.Text = Sheet_Name + " m:Z Self Impedance"	# Add Title
				#xl.ActiveChart.Legend.Delete()                                                	# Delete legend
				xl.ActiveChart.Axes(c.xlCategory).AxisTitle.Text = "Frequency in Hz"            # X Axis
				xl.ActiveChart.Axes(c.xlCategory).MinimumScale = 0                            	# Set minimum of x axis
				xl.ActiveChart.Axes(c.xlCategory).MaximumScale = scale_end                  	# Set maximum of x axis
				xl.ActiveChart.Axes(c.xlCategory).TickLabels.NumberFormat = "0"               	# Set number of decimals 0.0                                               
				xl.ActiveChart.Axes(c.xlValue).AxisTitle.Text = "Impedance in Ohms"      		# Y Axis
				xl.ActiveChart.Axes(c.xlValue).TickLabels.NumberFormat = "0"                   	# Set number of decimals  
			
			t2 = time.clock() - t1
			print1(1,"Graphing Z self impedance data, time taken: " + str(round(t2,2)) + " seconds",0)
			t1 = time.clock()

		if Excel_Export_Z12 == True:	# Export Z12 data
			# Insert Mutual Z_12 data to excel______________________________________________________________________________________________
			print1(1,"Inserting Z_12 data",0)
			if Excel_Export_RX == True or Excel_Export_Z == True:
				newcol = newcol + 1
			Z_12_First = newcol
			for x in FS_Results:													 
				if x[1] == "c:Z_12":
					ws.Range(ws.Cells(startrow-1,newcol),ws.Cells(endrow,newcol)).Value = list(zip(*[x])) 
					newcol = newcol + 1	
			Z_12_Last = newcol - 1
			t2 = time.clock() - t1
			print1(1,"Exporting Z_12 data self impedance data, time taken: " + str(round(t2,2)) + " seconds",0)	
			t1 = time.clock()    
	wb.Save()							# Save Workbook
	
	if HRM_Sim == True:
		HRM_endrow = startrow + len(HRM_Results[0]) - 1
		if Excel_Export_HRM == True:
			# Harmonic data to excel_________________________________________________________________________________________________________
			print1(1,"Inserting Harmonic data",0)
			if Excel_Export_RX == True or Excel_Export_Z == True or Excel_Export_Z12 == True:						# Adds a space between FS & harmonic data
				newcol = newcol + 1
			HRM_First = newcol
			HRM_scale = HRM_Results[0]
			HRM_scale1 = [int(int(x) / 50) for x in HRM_scale[4:]]
			HRM_scale = HRM_scale[:4]
			HRM_scale.extend(HRM_scale1)
			ws.Range(ws.Cells(startrow,newcol),ws.Cells(HRM_endrow,newcol)).Value = list(zip(*[HRM_scale]))	# Exports the Scale to excel
			newcol = newcol + 1	
			HRM_Base_Case_Pos = []
			for x in HRM_Results:																					# Exports the results to excel				 
				if x[0] == "m:HD":
					ws.Range(ws.Cells(startrow,newcol),ws.Cells(HRM_endrow,newcol)).Value = list(zip(*[x])) 
					if x[2] == "Base_Case":
						HRM_Base_Case_Pos.append([newcol])
					newcol = newcol + 1	
			HRM_Last = newcol - 1
			
			iec_limits = [["IEC", "61000-3-6", "Harmonics", "THD", 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 
							21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40],
							["IEC", "61000-3-6", "Limits", 3, 1.4, 2, 0.8, 2, 0.4, 2, 0.4, 1, 0.35, 1.5, 0.32, 1.5, 0.3, 0.3, 0.28, 1.2, 0.265, 0.93, 0.255, 0.2, 0.246, 0.88, 
							0.24, 0.816, 0.233, 0.2, 0.227, 0.703, 0.223, 0.66, 0.219, 0.2, 0.2158, 0.58, 0.2127, 0.55, 0.21, 0.2, 0.2075]]
			limits = iec_limits[1]
			limits = list(zip(*[iec_limits[1]]))
			
			# Graph Harmonic Distortion Charts
			if Excel_Export_Z == True:
				hrm_top = 500
			else:
				hrm_top = 45
			
			if len(HRM_Base_Case_Pos) > 1:			# If there is more than 1 Base Case then plot all the bases on one graph and then each base against its N-1 across else just plot them all on one graph.	
				hrm_no_of_contingencies = int(HRM_Base_Case_Pos[1][0]) - int(HRM_Base_Case_Pos[0][0])
				ws.Range(ws.Cells(1,1),ws.Cells(1,2)).Select()
				ch = ws.Shapes.AddChart(c.xlColumnClustered,30,hrm_top,825,400).Select()    						# AddChart(Type, Left, Top, Width, Height)
				xl.ActiveChart.ApplyLayout(9)																	# Select Layout 1-11
				xl.ActiveChart.ChartTitle.Characters.Text = Sheet_Name + " Base Case Harmonic Emissions v IEC Limits"		# Add Title
				xl.ActiveChart.SeriesCollection(1).Delete()
				#xl.ActiveChart.Legend.Delete()                                                					# Delete legend
				xl.ActiveChart.Axes(c.xlValue).AxisTitle.Text = "HD %"      									# Y Axis
				xl.ActiveChart.Axes(c.xlValue).TickLabels.NumberFormat = "0.0"                 					# Set number of decimals 
				xl.ActiveChart.Axes(c.xlCategory).AxisTitle.Text = "Harmonic"            						# X Axis
				#xl.ActiveChart.Axes(c.xlCategory).MinimumScale = 0                            					# Set minimum of x axis
				#xl.ActiveChart.Axes(c.xlCategory).TickLabels.NumberFormat = "0"               					# Set number of decimals 0.0                                               
				xl.ActiveChart.XValues = ws.Range(ws.Cells((startrow + 3), HRM_First),ws.Cells((HRM_endrow), HRM_First))			# X Value
				for hrm_col in HRM_Base_Case_Pos:			
					series = xl.ActiveChart.SeriesCollection().NewSeries()
					series_name1 = ws.Range(ws.Cells((startrow + 1), hrm_col[0]),ws.Cells((startrow + 2), hrm_col[0])).Value
					series_name = str(series_name1[0][0]) + "_" + str(series_name1[1][0])
					series.Values = ws.Range(ws.Cells((startrow + 3), hrm_col[0]),ws.Cells((HRM_endrow), hrm_col[0]))						# Y Value
					series.XValues = ws.Range(ws.Cells((startrow + 3), HRM_First),ws.Cells((HRM_endrow), HRM_First))
					series.Name = series_name																						#
				ws.Range(ws.Cells(startrow,newcol),ws.Cells(startrow + len(limits) - 1,newcol)).Value = limits						# Export the limits as far as the 40th Harmonic
				series = xl.ActiveChart.SeriesCollection().NewSeries()																# Add series to the graph
				series.Values = ws.Range(ws.Cells(startrow+3,newcol),ws.Cells(startrow + len(limits) - 1,newcol))					# Y Value
				series.XValues = ws.Range(ws.Cells((startrow + 3), HRM_First),ws.Cells((HRM_endrow), HRM_First))
				series.Name = "IEC 61000-3-6"																						# Name
				series.Format.Fill.Visible = True 												# Add fill to chart		
				series.Format.Fill.ForeColor.RGB = 12611584										# Colour for fill (red + green*256 + blue*256*256)
				series.Format.Fill.ForeColor.Brightness = 0.75									# Fill Brightness
				series.Format.Fill.Transparency = 0.75											# Fill Transparency
				series.Format.Fill.Solid														# Solid Fill
				series.Border.Color = 12611584													# Fill Colour
				series.Format.Line.Visible = True												# Series line is visible
				series.Format.Line.Weight = 1.5													# Set line weight for series
				series.Format.Line.ForeColor.RGB = 12611584										# Colour for line (red + green*256 + blue*256*256)
				series.AxisGroup = 2															# Move to Secondary Axes
				xl.ActiveChart.ChartGroups(2).Overlap = 100										# Edit Secondary Axis Overlap of bars
				xl.ActiveChart.ChartGroups(2).GapWidth = 0										# Edit Secondary Axis width between bars
				xl.ActiveChart.Axes(c.xlValue).MaximumScale = 3.5                               # Set scale Max
				xl.ActiveChart.Axes(c.xlValue, c.xlSecondary).MaximumScale = 3.5                # Set scale Min
				
				hrmb_count = 1
				for hrm_col in HRM_Base_Case_Pos:
					ws.Range(ws.Cells(1,1),ws.Cells(1,2)).Select()
					ch = ws.Shapes.AddChart(c.xlColumnClustered,30 + hrmb_count * 855,hrm_top,825,400).Select()    						# AddChart(Type, Left, Top, Width, Height)
					xl.ActiveChart.ApplyLayout(9)																	# Select Layout 1-11
					series_name1 = ws.Range(ws.Cells((startrow + 1), hrm_col[0]),ws.Cells((startrow + 2), hrm_col[0])).Value
					series_name = str(series_name1[0][0])
					xl.ActiveChart.ChartTitle.Characters.Text = Sheet_Name + " " + str(series_name) + " Harmonic Emissions v IEC Limits"	# Add Title
					xl.ActiveChart.SeriesCollection(1).Delete()
					#xl.ActiveChart.Legend.Delete()                                                					# Delete legend
					xl.ActiveChart.Axes(c.xlValue).AxisTitle.Text = "HD %"      									# Y Axis
					xl.ActiveChart.Axes(c.xlValue).TickLabels.NumberFormat = "0.0"                 					# Set number of decimals 
					xl.ActiveChart.Axes(c.xlCategory).AxisTitle.Text = "Harmonic"            						# X Axis
					#xl.ActiveChart.Axes(c.xlCategory).MinimumScale = 0                            					# Set minimum of x axis
					#xl.ActiveChart.Axes(c.xlCategory).TickLabels.NumberFormat = "0"               					# Set number of decimals 0.0                                               
					xl.ActiveChart.XValues = ws.Range(ws.Cells((startrow + 3), HRM_First),ws.Cells((HRM_endrow), HRM_First))			# X Value
					for hrm_col1 in list(range(hrm_col[0],(hrm_col[0] + hrm_no_of_contingencies))):			
						series = xl.ActiveChart.SeriesCollection().NewSeries()
						series_name1 = ws.Range(ws.Cells((startrow + 1), hrm_col1),ws.Cells((startrow + 2), hrm_col1)).Value
						series_name = str(series_name1[0][0]) + "_" + str(series_name1[1][0])
						series.Values = ws.Range(ws.Cells((startrow + 3), hrm_col1),ws.Cells((HRM_endrow), hrm_col1))						# Y Value
						series.XValues = ws.Range(ws.Cells((startrow + 3), HRM_First),ws.Cells((HRM_endrow), HRM_First))
						series.Name = series_name																						#
					ws.Range(ws.Cells(startrow,newcol),ws.Cells(startrow + len(limits) - 1,newcol)).Value = limits						# Export the limits as far as the 40th Harmonic
					series = xl.ActiveChart.SeriesCollection().NewSeries()																# Add series to the graph
					series.Values = ws.Range(ws.Cells(startrow+3,newcol),ws.Cells(startrow + len(limits) - 1,newcol))					# Y Value
					series.XValues = ws.Range(ws.Cells((startrow + 3), HRM_First),ws.Cells((HRM_endrow), HRM_First))
					series.Name = "IEC 61000-3-6"																						# Name
					series.Format.Fill.Visible = True 												# Add fill to chart		
					series.Format.Fill.ForeColor.RGB = 12611584										# Colour for fill (red + green*256 + blue*256*256)
					series.Format.Fill.ForeColor.Brightness = 0.75									# Fill Brightness
					series.Format.Fill.Transparency = 0.75											# Fill Transparency
					series.Format.Fill.Solid														# Solid Fill
					series.Border.Color = 12611584													# Fill Colour
					series.Format.Line.Visible = True												# Series line is visible
					series.Format.Line.Weight = 1.5													# Set line weight for series
					series.Format.Line.ForeColor.RGB = 12611584										# Colour for line (red + green*256 + blue*256*256)
					series.AxisGroup = 2															# Move to Secondary Axes
					xl.ActiveChart.ChartGroups(2).Overlap = 100										# Edit Secondary Axis Overlap of bars
					xl.ActiveChart.ChartGroups(2).GapWidth = 0										# Edit Secondary Axis width between bars
					xl.ActiveChart.Axes(c.xlValue).MaximumScale = 3.5                               # Set scale Max
					xl.ActiveChart.Axes(c.xlValue, c.xlSecondary).MaximumScale = 3.5                # Set scale Min
					hrmb_count = hrmb_count + 1
			else:
				ws.Range(ws.Cells(1,1),ws.Cells(1,2)).Select()
				ch = ws.Shapes.AddChart(c.xlColumnClustered,30,hrm_top,825,400).Select()    						# AddChart(Type, Left, Top, Width, Height)
				xl.ActiveChart.ApplyLayout(9)																	# Select Layout 1-11
				xl.ActiveChart.ChartTitle.Characters.Text = Sheet_Name + " Harmonic Emissions v IEC Limits"		# Add Title
				xl.ActiveChart.SeriesCollection(1).Delete()
				#xl.ActiveChart.Legend.Delete()                                                					# Delete legend
				xl.ActiveChart.Axes(c.xlValue).AxisTitle.Text = "HD %"      									# Y Axis
				xl.ActiveChart.Axes(c.xlValue).TickLabels.NumberFormat = "0.0"                 					# Set number of decimals 
				xl.ActiveChart.Axes(c.xlCategory).AxisTitle.Text = "Harmonic"            						# X Axis
				#xl.ActiveChart.Axes(c.xlCategory).MinimumScale = 0                            					# Set minimum of x axis
				#xl.ActiveChart.Axes(c.xlCategory).TickLabels.NumberFormat = "0"               					# Set number of decimals 0.0                                               
				xl.ActiveChart.XValues = ws.Range(ws.Cells((startrow + 3), HRM_First),ws.Cells((HRM_endrow), HRM_First))			# X Value
				for hrm_col in range(HRM_First + 1, HRM_Last + 1):			
					series_name1 = ws.Range(ws.Cells((startrow + 1), hrm_col),ws.Cells((startrow + 2), hrm_col)).Value
					series_name = str(series_name1[0][0]) + "_" + str(series_name1[1][0])
					series = xl.ActiveChart.SeriesCollection().NewSeries()
					series.Values = ws.Range(ws.Cells((startrow + 3), hrm_col),ws.Cells((HRM_endrow), hrm_col))						# Y Value
					series.XValues = ws.Range(ws.Cells((startrow + 3), HRM_First),ws.Cells((HRM_endrow), HRM_First))
					series.Name = series_name																						#
				ws.Range(ws.Cells(startrow,newcol),ws.Cells(startrow + len(limits) - 1,newcol)).Value = limits						# Export the limits as far as the 40th Harmonic
				series = xl.ActiveChart.SeriesCollection().NewSeries()																# Add series to the graph
				series.Values = ws.Range(ws.Cells(startrow+3,newcol),ws.Cells(startrow + len(limits) - 1,newcol))					# Y Value
				series.XValues = ws.Range(ws.Cells((startrow + 3), HRM_First),ws.Cells((HRM_endrow), HRM_First))
				series.Name = "IEC 61000-3-6"																						# Name
				series.Format.Fill.Visible = True 												# Add fill to chart		
				series.Format.Fill.ForeColor.RGB = 12611584										# Colour for fill (red + green*256 + blue*256*256)
				series.Format.Fill.ForeColor.Brightness = 0.75									# Fill Brightness
				series.Format.Fill.Transparency = 0.75											# Fill Transparency
				series.Format.Fill.Solid														# Solid Fill
				series.Border.Color = 12611584													# Fill Colour
				series.Format.Line.Visible = True												# Series line is visible
				series.Format.Line.Weight = 1.5													# Set line weight for series
				series.Format.Line.ForeColor.RGB = 12611584										# Colour for line (red + green*256 + blue*256*256)
				series.AxisGroup = 2															# Move to Secondary Axes
				xl.ActiveChart.ChartGroups(2).Overlap = 100										# Edit Secondary Axis Overlap of bars
				xl.ActiveChart.ChartGroups(2).GapWidth = 0										# Edit Secondary Axis width between bars
				xl.ActiveChart.Axes(c.xlValue).MaximumScale = 3.5                               # Set scale Max
				xl.ActiveChart.Axes(c.xlValue, c.xlSecondary).MaximumScale = 3.5                # Set scale Min
			
			t2 = time.clock() - t1
			print1(1,"Exporting Harmonic data, time taken: " + str(round(t2,2)) + " seconds",0)	
			t1 = time.clock()		
	
	#Position = [list(map(lambda xq: xq + startrow, scale_list_int)), list(range(R_First, R_Last)), list(range(R_First, R_Last)), list(range(X_First, X_Last)), 
	#			list(range(Z_First, Z_Last)), list(range(Z_12_First, Z_12_Last)), list(range(startrow, HRM_endrow)), list(range(HRM_First, HRM_Last))]
	#for ssert in Position:
		#print2(ssert)	    
	wb.Save()							# Save Workbook
	return

def Create_Textfile_Sheet(Sheet_Name, Text, wb):      # Extract information from out file
	t1 = time.clock()
	xl = win32com.client.gencache.EnsureDispatch('Excel.Application')   # Call disptach excel application excel
	c = win32com.client.constants                                       #
	ws = wb.Worksheets.Add()                                            # Add worksheet      	
	count  = 2
	for line in Text:
		ws.Cells(count,1).Value = line
		count = count + 1
	ws.Name = Sheet_Name                                                                                # Rename worksheet          
	wb.Save()
	t2 = time.clock() - t1
	print1(2,"Creating Sheet: " + Sheet_Name + " " + str(round(t2,2)) + " seconds",0)
	return
	
def Close_Workbook(wb,workbookname):		# Save and close Workbook
	print1(1,"Closing and Saving Workbook: " + workbookname,0)
	xl = win32com.client.gencache.EnsureDispatch('Excel.Application')   # Call disptach excel application excel
	wb.SaveAs(workbookname)                                             # Save Workbook"""
	wb.Close()                                                          # Close Workbook
	xl.Application.Quit()                                               # Quit Excel
	return

def ConvexHull1(pointlist):			# Gets the convex hull of a numpy array (if you have a list of tuples us np.array(pointlist) to convert
	R, X, Convex_Points = [], [], []
	cv = ConvexHull(pointlist)
	#R = pointlist[cv.vertices,0]									# the vertices of the convex hull
	#X = pointlist[cv.vertices,1]
	for x in cv.vertices:
		R.append(float(pointlist[x,0]))								# Converts the numpy floats back to regular float and attach
		X.append(float(pointlist[x,1]))
		#app.PrintPlain(pointlist[x,0])
	R.append(R[0])
	X.append(X[0])
	Convex_Points.append(R)
	Convex_Points.append(X)
	return Convex_Points
	
# Main Engine --------------------------------------------------------------------------------------------------------------------------------	
# --------------------------------------------------------------------------------------------------------------------------------------------	
Error_Count = 1

# User Input (All info is checked to see if it exists in the case file
# -------------------------------------------------------------------------------------------------------------------------------
"""Import_Workbook = filelocation + "Harmonic_Inputs.xlsx"
Results_Export_Folder = "C:\\Users\\oconnell_b\\Desktop\\Scrap\\"			# Folder to Export Excel Results too
Excel_Results = Results_Export_Folder + "Results_" + start1					# Name of Exported Results File
Progress_Log = Results_Export_Folder + "Progress_Log_" + start1 + ".txt"	# Progress File
Error_Log = Results_Export_Folder + "Error_Log_" + start1 + ".txt"			# Error File
Random_Log = Results_Export_Folder + "Random_Log_" + start1 + ".txt"		# For printing random info solely for development 
Net_Elm = "EIRGRID.ElmNet"					# Where all the Network elements are stored
Mut_Elm_Fld = "ElmMut" + start1				# Name of the folder to create under the network elements to store mutual impedances
Results_Folder = "Results_" + start1		# Name of the folder to keep results under studycase
Operation_Scenario_Folder = "Op_Scenarios_"	+ start1 # Name of the folder to store Operational Scenarios
Pre_Case_Check = False					# Checks the N-1 for load flow convergence and saves operational scenarios.
FS_Sim = True						# Carries out Frequency Sweep Analysis
HRM_Sim = True						# Carries out Harmonic Load Flow Analysis
Plot_in_PF = False						# Plot the Frequency Sweeps in PF **************Currently not working
Skip_Unsolved_Ldf = False				# Skips the frequency sweep if the load flow doesn't solve
Delete_Created_Folders = False			# Deletes the Results folder, Mutual Elements and the Operational Scenario folder
Export_to_Excel = True					# Export the results to Excel
Excel_Visible = False					# Makes Excel Visible while plotting, Can be annoying if you are doing other work as if you click the excel screen it stops the simulation
Excel_Export_RX = True					# Export RX and graph the Impedance Loci in Excel
Excel_Convex_Hull = True				# This calculates the minimum points for the Loci
Excel_Export_Z = False					# Graph the Frequency Sweeps in Excel
Excel_Export_Z12 = False				# Export Mutual Impedances to excel
Excel_Export_HRM = False				# Export Harmonic Data

# List of Study case & Scenario to start with [Name, location\\studycase, location\\scenario]
List_of_Studycases = [["SV Base Case","Barry\\Summer Valley Base Case.IntCase", "Summer Valley\\2014 Summer Valley_base.IntScenario"],
					["SV with Glen","Barry\\Summer Valley Base Case with Glen.IntCase", "Summer Valley\\2014 Summer Valley_base.IntScenario"],
					["SV with T144","Barry\\Summer Valley Base Case with T144.IntCase", "Summer Valley\\2014 Summer Valley_base.IntScenario"],	
					["SV Both","Barry\\Summer Valley Base Case with Both.IntCase", "Summer Valley\\2014 Summer Valley_base.IntScenario"],
					["WP Base Case","Barry\\Winter Peak Base Case.IntCase", "Winter Peak\\2014 Winter Peak Base.IntScenario"],
					["WP with Glen","Barry\\Winter Peak Base Case with Glen.IntCase", "Winter Peak\\2014 Winter Peak Base.IntScenario"],
					["WP with T144","Barry\\Winter Peak Base Case with T144.IntCase", "Winter Peak\\2014 Winter Peak Base.IntScenario"],
					["WP Both","Barry\\Winter Peak Base Case with Both.IntCase", "Winter Peak\\2014 Winter Peak Base.IntScenario"]]			
			
# What Terminals do you want to look at (must be a minimum of 2 terminals specified)
List_of_Points = [("Cauteen.ElmSubstat", "110 kV A1.ElmTerm"), ("Doon.ElmSubstat","110 kV A1.ElmTerm")] 						

# List of N-1 Contingencies (must be a minimum of 2 contingencies) 
## -------------------------------------------------------------------------------------
# Notes Leave item 1 Base Case untouched 
# Enter the name and as many contingencies you want to switch ("name_of_the_contingency_as_one_word", "Name_of_Element_1_to_switch", "Name_of_Element_2_to_switch")

List_of_Contingencies = [("Base_Case",0),
						("KIL_CHA",["Killonan", "110 kV Charleville CB.ElmCoup"],["Charleville", "110 kV Killonan CB.ElmCoup"]),
						("KIL_LIM1",["Killonan", "110 kV Limerick #1 CB.ElmCoup"],["Limerick", "110 kV Killonan #1 CB.ElmCoup"]),
						("KIL_LIM2",["Killonan", "110 kV Limerick #2 CB.ElmCoup"],["Limerick", "110 kV Killonan #2 CB.ElmCoup"]),
						("KIL_SING",["Killonan", "110 kV Singland CB.ElmCoup"],["Singland", "110 kV Killonan CB.ElmCoup"]),						
						("KIL_TA",["Killonan", "220 kV Tarbert CB.ElmCoup"],["Tarbert", "220 kV Killonan CB.ElmCoup"]),	
						("BDN_CUL",["Ballydine", "110 kV Cullenagh CB.ElmCoup"],["Cullenagh", "110 kV Ballydine CB.ElmCoup"]),	
						("CAH_BAR_KRA",["Knockraha", "110 kV Cahir CB.ElmCoup"],["Cahir", "110 kV Barrymore CB.ElmCoup"]),	
						("CAH_DOO",["Doon", "110 kV Cahir CB.ElmCoup"],["Cahir", "110 kV Doon CB.ElmCoup"]),	
						("CAH_KIL",["Kill Hill", "110 kV Cullenagh CB.ElmCoup"],["Cahir", "110 kV Thurles CB.ElmCoup"]),	
						("CTN_KIL",["Cauteen", "110 kV Killonan CB.ElmCoup"],["Killonan", "110 kV Cauteen CB.ElmCoup"]),
						("CTN_TIP",["Cauteen", "110 kV Tipperary CB.ElmCoup"],["Tipperary", "110 kV Cauteen CB.ElmCoup"]),
						("CUL T2101",["Cullenagh", "110 kV T2101 CB.ElmCoup"],["Cullenagh", "220 kV T2101 CB.ElmCoup"]),
						("CUL_KRA",["Cullenagh", "220 kV Knockraha CB.ElmCoup"],["Knockraha", "220 kV Cullenagh CB.ElmCoup"]),	
						("DOO_BDN",["Doon", "110 kV Ballydine CB.ElmCoup"],["Ballydine", "110 kV Doon CB.ElmCoup"]),		
						("KHIL_THU",["Kill Hill", "110 kV Doon CB.ElmCoup"],["Thurles", "110 kV Cahir CB.ElmCoup"]),						
						("KIL T2101",["Killonan", "110 kV T2101 CB.ElmCoup"],["Killonan", "220 kV T2101 CB.ElmCoup"]),
						("KIL_SHA",["Killonan", "220 kV Shannonbridge CB.ElmCoup"],["Shannonbridge", "220 kV Killonan CB.ElmCoup"]),	
						("KRA T2101",["Knockraha", "110 kV T2101 CB.ElmCoup"],["Knockraha", "220 kV T2101 CB.ElmCoup"]),
						("KIL_KRA",["Killonan", "220 kV Knockraha CB.ElmCoup"],["Knockraha", "220 kV Killonan CB.ElmCoup"]),	
						("SHA T2101",["Shannonbridge", "110 kV T2101 CB.ElmCoup"],["Shannonbridge", "220 kV T2101 CB.ElmCoup"]),
						("THU_IKN_SH",["Shannonbridge", "110 kV Thurles CB.ElmCoup"],["Thurles", "110 kV Ikerrin CB.ElmCoup"]),
						("TIP_CAH",["Cahir", "110 kV Tipperary CB.ElmCoup"],["Tipperary", "110 kV Cahir CB.ElmCoup"]),						
						]	
"""

# Enter what Variables you want to look at for terminals
FS_Terminal_Variables = ["m:R", "m:X", "m:Z", "m:phiz"]	
Mutual_Variables = ["c:Z_12"]
HRM_Terminal_Variables = ["m:HD"]
# Import Excel
Import_Workbook = filelocation + "HAST_V1_2_Inputs.xlsx"					# Gets the CWD current working directory
Variation_Name = "Temporary_Variation" + start1
analysis_dict = Import_Excel_Harmonic_Inputs(Import_Workbook) 			# Reads in the Settings from the spreadsheet
Study_Settings = analysis_dict["Study_Settings"]						
if len(Study_Settings) != 20:
	print2("Error Check input Study Settings there should be 20 Items in the list there are only: " + len(Study_Settings) + " " + Study_Settings)
if Study_Settings[0] == None:											# If there is no output location in the spreadsheet it sets it to the CWD
	Results_Export_Folder = filelocation
else:
	Results_Export_Folder = Study_Settings[0]								# Folder to Export Excel Results too
Excel_Results = Results_Export_Folder + Study_Settings[1] + start1			# Name of Exported Results File
Progress_Log = Results_Export_Folder + Study_Settings[2] + start1 + ".txt"	# Progress File
Error_Log = Results_Export_Folder + Study_Settings[3] + start1 + ".txt"		# Error File
Random_Log = Results_Export_Folder + "Random_Log_" + start1 + ".txt"		# For printing random info solely for development 
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
print1(1,Title,0)
for keys,values in analysis_dict.items():									# Prints all the inputs to progress log
	print1(1, keys, 0)
	print1(1, values, 0)
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
	print2("Error - Check excel input Loadflow_Settings there should be 55 Items in the list there are only: " + len(Load_Flow_Setting) + " " + Load_Flow_Setting)
Fsweep_Settings = analysis_dict["Frequency_Sweep"]							# Imports Settings for Frequency Sweep calculation
if len(Fsweep_Settings) != 16:												# Check there are the right number of inputs
	print2("Error - Check excel input Frequency_Sweep there should be 16 Items in the list there are only: " + len(Fsweep_Settings) + " " + Fsweep_Settings)
Harmonic_Loadflow_Settings = analysis_dict["Harmonic_Loadflow"]				# Imports Settings for Harmonic LDF calculation
if len(Harmonic_Loadflow_Settings) != 15:									# Check there are the right number of inputs
	print2("Error - Check excel input Harmonic_Loadflow Settings there should be 17 Items in the list there are only: " + len(Harmonic_Loadflow_Settings) + " " + Harmonic_Loadflow_Settings)


List_of_Studycases1 = Check_List_of_Studycases(List_of_Studycases)			# This loops through all the studycases and operational scenarios listed and checks them skips any ones which don't solve

if FS_Sim == True or HRM_Sim == True:
	FS_Contingency_Results, HRM_Contingency_Results = [], []
	count_studycase = 0
	while count_studycase < len(List_of_Studycases1):												# Loop Through (Study Cases, Operational Scenarios)
		prj = ActivateProject(List_of_Studycases1[count_studycase][1])								# Activate Project
		if len(str(prj)) > 0:
			StudyCase, study_error = ActivateStudyCase(List_of_Studycases1[count_studycase][2])		# Activate Study Case in List
			Scenario, scen_error = ActivateScenario(List_of_Studycases1[count_studycase][3])		# Activate corresponding operational Scenario			
			Study_Case_Folder = app.GetProjectFolder("study")										# Returns string the location of the project folder for "study", (Ops) "scen" , "scheme" (Variations) Python reference guide 4.6.19 IntPrjfolder
			Operation_Case_Folder = app.GetProjectFolder("scen")						
			Variations_Folder = app.GetProjectFolder("scheme")
			Active_variations = GetActiveVariations()
			Variation = Create_Variation(Variations_Folder,"IntScheme",Variation_Name)
			ActivateVariation(Variation)
			Stage = Create_Stage(Variation,"IntSstage",Variation_Name)
			New_Contingency_List, Con_ok = Check_Contingencis(List_of_Contingencies) 				# Checks to see all the elements in the contingency list are in the case file
			Terminals_index, Term_ok = Check_Terminals(List_of_Points)								# Checks to see if all the terminals are in the case file skips any that aren't
			studycase_results_folder, folder_exists1 = Create_Folder(StudyCase, Results_Folder)					
			op_sc_results_folder, folder_exists2 = Create_Folder(Operation_Case_Folder, Operation_Scenario_Folder)
			Net_Elm1 = Get_Object(Net_Elm)															# Gets the Network Elements ElmNet folder
			if len(Net_Elm1) < 1:
				print2("Could not find Network Element folder, Note: this is case sensitive :" + str(Net_Elm))
			if len(Terminals_index) > 1 and Excel_Export_Z12 == True:
				studycase_mutual_folder, folder_exists3 = Create_Folder(Net_Elm1[0], Mut_Elm_Fld)		# Create Folder for Mutual Elements
				List_of_Mutual = Create_Mutual_Impedance_List(studycase_mutual_folder, Terminals_index)	# Create List of mutual impedances between the terminals in the folder
			else:
				Excel_Export_Z12 = False															# Can't Export mutual impedances if you give it only one bus
			count = 0
			while count < len(New_Contingency_List):												# Loop Through Contingency list						
				print1(2,"Carrying out Contingency: " + New_Contingency_List[count][0],0)
				DeactivateScenario()																# Can't copy activated Scenario so deactivate it
				object_exists, new_object = Check_If_Object_Exists(op_sc_results_folder, List_of_Studycases1[count_studycase][0] + str("_" + New_Contingency_List[count][0] + ".IntScenario"))
				if object_exists == 0:
					new_scenario = Add_Copy(op_sc_results_folder,Scenario,List_of_Studycases1[count_studycase][0] + str("_" + New_Contingency_List[count][0]))	# Copies the base scenario	
				else:
					new_scenario = new_object[0]
				scen_error = ActivateScenario1(new_scenario)										# Activates the base scenario
				if New_Contingency_List[count][0] != "Base_Case":									# Apply Contingencies if it is not the base case
					for switch in New_Contingency_List[count][1:]:																								
						Switch_Coup(switch[0],switch[1])
				SaveActiveScenario()
				if FS_Sim == True:																	# Skips the Frequency Analysis
					sweep = Create_Results_File(studycase_results_folder, New_Contingency_List[count][0] + "_FS",9)		# Create Results File
					trm_count = 0
					while trm_count < len(Terminals_index):											# Add terminal variables to the Results file													
						Add_Vars_Res(sweep, Terminals_index[trm_count][3], FS_Terminal_Variables)
						trm_count = trm_count + 1
					if Excel_Export_Z12 == True:
						for mut in List_of_Mutual:													# Adds the mutual impedance data to Results File
							Add_Vars_Res(sweep, mut[2], Mutual_Variables)
					Fsweep_err_cde = FSweep(sweep,Fsweep_Settings)									# Carry out Frequency Sweep
					if Fsweep_err_cde == 0:															# Skips the contingency if Frequency Sweep doesn't solve
						fs_scale, fs_res = Retrieve_Results(sweep,0)		
						fs_scale.insert(1,"Frequency in Hz")										# Arranges the Frequency Scale
						fs_scale.insert(1,"Scale")
						fs_scale.pop(3)
						for tope in fs_res:															# Adds the additional information to the results file
							tope.insert(1,New_Contingency_List[count][0])							# Op scenario
							tope.insert(1,List_of_Studycases1[count_studycase][0])					# Study case description
							FS_Contingency_Results.append(tope)										# Results
					else:
						print2("Error with Frequency Sweep Simulation: " + List_of_Studycases1[count_studycase][0] + New_Contingency_List[count][0])
				else:
					fs_scale = []
				if HRM_Sim == True:				
					harm = Create_Results_File(studycase_results_folder, New_Contingency_List[count][0] + "_HLF",6)		# Creates the Harmonic Results File
					trm_count = 0
					while trm_count < len(Terminals_index):											# Add terminal variables to the Results file													
						Add_Vars_Res(harm, Terminals_index[trm_count][3], HRM_Terminal_Variables)
						trm_count = trm_count + 1
					Harm_err_cde = HarmLoadFlow(harm,Harmonic_Loadflow_Settings)
					if Harm_err_cde == 0:
						hrm_scale, hrm_res = Retrieve_Results(harm,1)
						hrm_scale.insert(1,"THD")													# Inserts the THD
						hrm_scale.insert(1,"Harmonic")												# Arranges the Harmonic Scale
						hrm_scale.insert(1,"Scale")
						hrm_scale.pop(4)															# Takes out the 50 Hz
						hrm_scale.pop(4)		
						for res12 in hrm_res:
							thd1 = re.split(r'[\\.]',res12[1])
							thd2 = app.GetCalcRelevantObjects(thd1[11] + ".ElmSubstat")
							thdz = False
							if thd2[0] != None:
								thd3 = thd2[0].GetContents()
								for thd4 in thd3:
									if (thd1[13] + ".ElmTerm") in str(thd4):
										THD = thd4.GetAttribute('m:THD')
										thdz = True
							elif thd2[0] != None or thdz == False:
								THD = "NA"
							#thd4 = app.SearchObjectByForeignKey(thd1[11] + ".ElmSubstat")
							res12.insert(2,THD)														# Insert THD
							res12.insert(2,New_Contingency_List[count][0])							# Op scenario
							res12.insert(2,List_of_Studycases1[count_studycase][0])					# Study case description
							res12.pop(5)
							HRM_Contingency_Results.append(res12)									# Results
					else:
						print2("Error with Harmonic Simulation: " + List_of_Studycases1[count_studycase][0] + New_Contingency_List[count][0])
				else:
					hrm_scale =[]
				count = count + 1
			print1(2,"",0)
			Scenario, scen_error = ActivateScenario(List_of_Studycases1[count_studycase][3])		# Activate the base case scenario this just ensures that when the script finishes using PF that it goes back to a regular case
			count_studycase = count_studycase + 1
			print1(2,"",0)
			if Delete_Created_Folders == True:														# Deletes Folder Created by automation
				Delete_Object(studycase_results_folder)						
				Delete_Object(op_sc_results_folder)	
				if Excel_Export_Z12 == True:
					Delete_Object(studycase_mutual_folder)
				Variation.Deactivate
				Delete_Object(Variation)
		else:
			print2("Could Not Activate Project: " + Project_Name)	

	if Export_to_Excel == True:																# This Exports the Results files to Excel in terminal format
		print1(1,"\nProcessing Results and output to Excel",0)
		start2 = time.clock()																# Used to calc the total excel export time
		wb = Create_Workbook(Excel_Results)													# Creates Workbook	
		trm1_count = 0
		while trm1_count < len(Terminals_index):											# For Terminals in the index loop through creating results to pass to excel sheet creator
			start3 = time.clock()															# Used for measuring time to create a sheet
			FS_Terminal_Results = []														# Creates a Temporary list to pass through terminal data to excel to create the terminal sheet
			if FS_Sim == True:
				start4 = time.clock()
				FS_Terminal_Results.append(fs_scale)										# Adds the scale to terminal
				for results34 in FS_Contingency_Results:									# Adds each contingency to the terminal results
					if str(Terminals_index[trm1_count][3]) == results34[3]:					# Checks it it the right terminal and adds it
						results34.pop(3)													# Takes out the terminal  PF object (big long string)
						FS_Terminal_Results.append(results34)								# Append terminal data to the results list to be later passed to excel
				#print1(1,"Process Results RX & Z in Python: " + str(round((time.clock() - start4),2)) + " Seconds",0)		# Returns python results processing time
				if Excel_Export_Z12 == True:
					start5 = time.clock()
					for results35 in FS_Contingency_Results:								# Adds each contingency to the terminal results			
						for tgb in List_of_Mutual:
							if Terminals_index[trm1_count][3] == tgb[3]:
								if str(tgb[2]) == str(results35[3]):						# Checks it it the right terminal and adds it
									results35.pop(3)										# Takes out the terminal  PF object (big long string)
									results35.insert(0,tgb[1])								# Adds in the Mutual tag ie Letterkenny_Binbane
									FS_Terminal_Results.append(results35)					# If it is the right terminal append
					print1(1,"Process Results Z12 in Python: " + str(round((time.clock() - start5),2)) + " Seconds",0)		# Returns python results processing time
			HRM_Terminal_Results = []														# Creates a Temporary list to pass through terminal data to excel to create the terminal sheet
			if HRM_Sim == True:	
				start6 = time.clock()
				HRM_Terminal_Results.append(hrm_scale)										# Adds the scale to terminal
				if Excel_Export_HRM == True:
					for results35 in HRM_Contingency_Results:								# Adds each contingency to the terminal results
						if str(Terminals_index[trm1_count][3]) == results35[1]:				# Checks it it the right terminal and adds it
							results35.pop(1)												# Takes out the terminal  PF object (big long string)
							HRM_Terminal_Results.append(results35)							# Append terminal data to the results list to be later passed to excel				
				print1(1,"Process Results HRM in Python: " + str(round((time.clock() - start6),2)) + " Seconds",0)		# Returns python results processing time
			Create_Sheet_Plot(Terminals_index[trm1_count][0],FS_Terminal_Results, HRM_Terminal_Results, wb)				# Uses the terminal results to create a sheet and graph
			trm1_count = trm1_count + 1
		# progress_txt = ReadTextfile(Progress_Log)
		# Create_Textfile_Sheet("Progress_Log", progress_txt, wb)
		Close_Workbook(wb,Excel_Results)																# Closes and saves the workbook
		print1(2,"Total Excel Export Time: " + str(round((time.clock() - start2),2)) + " Seconds",0)	# Returns the Total Export time	

print1(2,"Total Time: " + str(round((time.clock() - start),2)) + " Seconds",0)							# Returns the Calc time

# End of Script
# --------------------------------------------------------------------------------------------------------------------------------------------

#case = app.GetActiveStudyCase()												# Get active study case
#app.PrintPlain(str(Study_Case_Folder))
#StudyCases = Study_Case_Folder.GetContents("*.IntCase")						# Gets the contents of the study case folder
#sad = Study_Case_Folder.GetContents("SNV Final Case - Max Gen with SVC No Fil\\2014 Summer Valley Base Case.IntCase")
#app.PrintPlain("sad")
#app.PrintPlain(sad)
#for i in StudyCases:
#	app.PrintPlain(i)
#allStudyCases = Study_Case_Folder.GetContents()
#for studyCase in allStudyCases:
#	app.PrintPlain(studyCase)
#	pp = studyCase.GetContents("*.IntCase")
#	app.PrintPlain(pp)
	#studyCase.Activate()

#						if Plot_in_PF == True:														# Plot in PF		
#							for var1 in Terminals_index:											# Add terminal variables to the Plot file
#								Plot(var1[0][:-11],'VisPlot',sweep, var1[2], "m:Z", New_Contingency_List[count][0], count)	
	
#sys.exit()
	
# Get Object Class 
#str = studyCase.GetClassName()
#app.PrintPlain(str)

# Export all plots available in a project	
#obj=app.GetGraphicsBoard()
#VIPages=obj.GetContents('*.SetVipage')
#for i in VIPages[0]:
#   obj.Show(i)
#   Page_name=i.loc_name
#   File_name=('D:\\Users\\PowerFactory\\%s' %(Page_name))
#   obj.WriteWMF(File_name)
	
# Text Control
#app.ClearOutputWindow()			# Clear Output Window
#app.PrintError(str message)	# Prints message as an error
#app.PrintInfo(str message)		# Prints message as info
#app.PrintPlain(str message)	# Prints message as plain
#app.PrintWarn(str message)		# Prints message as a warning

#Shc_folder = app.GetFromStudyCase('IntEvt');

#terminals = app.GetCalcRelevantObjects("*.ElmTerm")
#lines = app.GetCalcRelevantObjects("*.ElmTerm")
#syms = app.GetCalcRelevantObjects("*.ElmSym")

#Shc_folder.CreateObject('EvtSwitch', 'evento de generacion');
#EventSet = Shc_folder.GetContents();  
#evt = EventSet[0];

#evt.time =1.0

#evt.p_target = syms[1]

#for sym in syms:
#    elmres.AddVars(sym,'s:xspeed')

    
#ini.Execute()
#sim.Execute()

#evt.Delete()

#comres = app.GetFromStudyCase('ComRes'); 
#comres.iopt_csel = 0
#comres.iopt_tsel = 0
#comres.iopt_locn = 2
#comres.ciopt_head = 1
#comres.pResult=elmres
#comres.f_name = r'C:\Users\jmmauricio\hola.txt'
#comres.iopt_exp=4
#comres.Execute()


# Get lists of buses and lines
#buses = app.GetCalcRelevantObjects('*.ElmTerm')
#lines = app.GetCalcRelevantObjects('*.ElmLne')
#cnv_gens = app.GetCalcRelevantObjects('*.ElmSym')
#wind_gens = app.GetCalcRelevantObjects('*.ElmAsm')
#cnv_gens = app.GetCalcRelevantObjects('*.ElmSym')
#cnv_gens = app.GetCalcRelevantObjects('*.ElmSym')

# Print bus voltages
#for bus in buses:
    # Only consider busbars (iUsage = 0) and in-service buses
#    if bus.iUsage == 0 and bus.outserv == 0:
#        bus_v = round(bus.GetAttribute('m:u'),2)
#        app.PrintPlain('Voltage on bus ' + str(bus) + ': ' + str(bus_v) + 'pu')

# Print loading on lines
#for line in lines:
#    if line.outserv == 0:
#        loading = round(line.GetAttribute('m:loading'),2)
#        app.PrintPlain('Loading on line ' + str(line) + ': ' + str(loading) + '%%')
