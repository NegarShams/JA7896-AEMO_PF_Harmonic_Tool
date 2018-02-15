"""
	Classes containing details needed for pf extract
"""

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


class PFStudyCase:
	""" Class containing the details for each new study case created """
	def __init__(self, full_name, list_parameters, cont_name, sc, op, prj, res_folder, task_auto, uid):
		"""
			Initialises the class with a list of parameters taken from the Study Settings import
		:param str full_name:  Full name of study case continaining base case and contingency
		:param list list_parameters:  List of parameters from the Base_Scenaraios inputs sheet
		:param str cont_name:  Name of contingency being considered
		:param object sc:  Handle to newly created study case
		:param object op:  Handle to newly created operational scenario
		:param object prj:  Handle to project in which this study case is contained
		:param object res_folder:  Handle to the folder which will contain all the results created as part of this project
		:param object task_auto:  Handle to the Task Automation object created for this project studies
		:param string uid:  Unique identifier time added to new files created
		"""
		# Strings that are used to store
		self.name = full_name
		self.base_name = list_parameters[0]
		self.prj_name = list_parameters[1]
		self.sc_name = list_parameters[2]
		self.op_name = list_parameters[3]
		self.cont_name = cont_name
		self.uid = uid

		# Handle for study cases that will require activating
		self.sc = sc
		self.op = op
		self.prj = prj
		self.res_folder = res_folder
		self.task_auto = task_auto

		# Attributes set during study completion
		self.frq = None
		self.hldf = None

	def create_freq_sweep(self, results_file, settings):
		"""
			Create a frequency sweep command in the study_case and return this as a reference
		:param object results_file:  Reference to the power factory results file for frequency sweep results
		:param list settings:  Settings for the frequency sweep to be created
		:return object frq_sweep:  Handle to the frq_sweep command that has been created
		"""
		# Create a new frequency sweep command object and store it in the study case
		frq = create_object(self.sc, 'ComFsweep', 'FSweep_{}'.format(self.uid))

		## Frequency Sweep Settings
		## -------------------------------------------------------------------------------------
		# Basic
		# TODO: Check whether all settings from input file are actually used
		frq.iopt_net = settings[2]  # Network Representation (0=Balanced 1=Unbalanced)
		frq.fstart = settings[3]  # Impedance Calculation Start frequency
		frq.fstop = settings[4]  # Stop Frequency
		frq.fstep = settings[5]  # Step Size
		frq.i_adapt = settings[6]  # Automatic Step Size Adaption
		frq.frnom = settings[7]  # Nominal Frequency
		frq.fshow = settings[8]  # Output Frequency
		frq.ifshow = settings[9]  # Harmonic Order
		frq.p_resvar = results_file  # Results Variable
		# TODO: Load flow settings for frequency sweep are currently not configured
		# frq.cbutldf = fsweep_settings[11]                 # Load flow

		# Advanced
		frq.errmax = settings[12]  # Setting for Step Size Adaption    Maximum Prediction Error
		frq.errinc = settings[13]  # Minimum Prediction Error
		frq.ninc = settings[14]  # Step Size Increase Delay
		frq.ioutall = settings[15]  # Calculate R, X at output frequency for all nodes

		self.frq = frq
		return self.frq

	def create_harm_load_flow(self, results_file, settings):  # Inputs load flow settings and executes load flow
		"""
			Runs harmonic load flow
		:param object results_file: Results variable provided as an input to the powerfactory harmonic load flow
		:param list settings: Harmonic load flow settings
		:return object hldf:  Handle to the hldf that has just been created
		"""
		# Create a new harmonnic load flow object and store it in the study case
		hldf = create_object(self.sc, 'ComHldf', 'HLDF_{}'.format(self.uid))

		## Loadflow settings
		## -------------------------------------------------------------------------------------
		# Basic
		hldf.iopt_net = settings[0]  # Calculation method (0 Balanced AC, 1 Unbalanced AC, DC)
		hldf.iopt_allfrq = settings[1]  # Calculate Harmonic Load Flow 0 - Single Frequency 1 - All Frequencies
		hldf.iopt_flicker = settings[2]  # Calculate Flicker
		hldf.iopt_SkV = settings[3]  # Calculate Sk at Fundamental Frequency
		hldf.frnom = settings[4]  # Nominal Frequency
		hldf.fshow = settings[5]  # Output Frequency
		hldf.ifshow = settings[6]  # Harmonic Order
		hldf.p_resvar = results_file  # Results Variable
		# TODO: No settings are currently provided for the load flow parameters
		# hldf.cbutldf =  harmonic_loadflow_settings[8]               	# Load flow

		# IEC 61000-3-6
		hldf.iopt_harmsrc = settings[9]  # Treatment of Harmonic Sources

		# Advanced Options
		hldf.iopt_thd = settings[10]  # Calculate HD and THD 0 Based on Fundamental Frequency values 1 Based on rated voltage/current
		hldf.maxHrmOrder = settings[11]  # Max Harmonic order for calculation of THD and THF
		hldf.iopt_HF = settings[12]  # Calculate Harmonic Factor (HF)
		hldf.ioutall = settings[13]  # Calculate R, X at output frequency for all nodes
		hldf.expQ = settings[14]  # Calculation of Factor-K (BS 7821) for Transformers

		self.hldf = hldf
		return self.hldf

class PFProject:
	""" Class contains reference to a project, results folder and associated task automation file"""
	def __init__(self, name, prj, res_folder, task_auto):
		"""
			Initialise class
		:param str name:  project name
		:param object prj:  project reference
		:param object res_folder:  folder reference
		:param object task_auto:  task automation reference
		"""
		self.name = name
		self.prj = prj
		self.res_folder = res_folder
		self.task_auto = task_auto
		self.sc_cases = []