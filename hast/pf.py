"""
	Classes containing details needed for pf extract
"""

import math

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

def remove_string_endings(astring, trailing):
	"""
		Function to strip the end from a string if it exists, used to remove .IntCase
	:param str astring:  Initial string
	:param str trailing:  Trailing string to be removed if exists
	:return str astring:  String returned without trail if it has been removed
	"""
	if astring.endswith(trailing):
		return astring[:-len(trailing)]
	return astring


class PFStudyCase:
	""" Class containing the details for each new study case created """
	def __init__(self, full_name, list_parameters, cont_name, sc, op, prj, task_auto, uid):
		"""
			Initialises the class with a list of parameters taken from the Study Settings import
		:param str full_name:  Full name of study case continaining base case and contingency
		:param list list_parameters:  List of parameters from the Base_Scenaraios inputs sheet
		:param str cont_name:  Name of contingency being considered
		:param object sc:  Handle to newly created study case
		:param object op:  Handle to newly created operational scenario
		:param object prj:  Handle to project in which this study case is contained
		:param object task_auto:  Handle to the Task Automation object created for this project studies
		:param string uid:  Unique identifier time added to new files created
		"""
		# Strings that are used to store
		self.name = full_name
		self.base_name = list_parameters[0]
		self.prj_name = list_parameters[1]
		# #self.sc_name = list_parameters[2]
		self.sc_name = remove_string_endings(astring=list_parameters[2], trailing='.IntCase')
		self.op_name = remove_string_endings(astring=list_parameters[3], trailing='.IntScenario')
		self.cont_name = cont_name
		self.uid = uid

		# Handle for study cases that will require activating
		self.sc = sc
		self.op = op
		self.prj = prj
		self.task_auto = task_auto

		# Attributes set during study completion
		self.frq = None
		self.hldf = None
		self.fs_results = None
		self.hldf_results = None
		self.fs_scale = []
		self.hrm_scale = []

		# Disctionary for looking up frequency scan results
		self.fs_res = dict()

	def create_freq_sweep(self, results_file, settings):
		"""
			Create a frequency sweep command in the study_case and return this as a reference
		:param object results_file:  Reference to the power factory results file for frequency sweep results
		:param list settings:  Settings for the frequency sweep to be created
		:return object frq_sweep:  Handle to the frq_sweep command that has been created
		"""
		self.fs_results = results_file
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
		self.hldf_results = results_file
		# Create a new harmonic load flow object and store it in the study case
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

	def process_fs_results(self):
		"""
			Function extracts and processes the load flow results for this study case
		:return list fs_res
		"""
		fs_scale, fs_res = retrieve_results(self.fs_results, 0)
		fs_scale.insert(1,"Frequency in Hz")										# Arranges the Frequency Scale
		fs_scale.insert(1,"Scale")
		fs_scale.pop(3)
		for tope in fs_res:															# Adds the additional information to the results file
			# #tope.insert(1, New_Contingency_List[count][0])							# Op scenario
			tope.insert(1, self.cont_name)											# Contingency name
			# #tope.insert(1,List_of_Studycases1[count_studycase][0])					# Study case description
			# # tope.insert(1, self.sc_name)											# Study case description
			# Added in to include scenario name as well
			# # tope.insert(1, '{}_{}'.format(self.sc_name, self.op_name))  # Study case description
			# Using base_name as description of study_case
			tope.insert(1, self.base_name)  # Study case description

		self.fs_scale = fs_scale

		return fs_res

	def process_hrlf_results(self, logger):
		"""
			Process the hrlf results ready for inclusion into spreadsheet
		:return hrm_res
		"""
		hrm_scale, hrm_res = retrieve_results(self.hldf_results, 1)

		hrm_scale.insert(1,"THD")													# Inserts the THD
		hrm_scale.insert(1,"Harmonic")												# Arranges the Harmonic Scale
		hrm_scale.insert(1,"Scale")
		hrm_scale.pop(4)															# Takes out the 50 Hz
		hrm_scale.pop(4)
		for res12 in hrm_res:
			# Rather than retrieving THD from the calculated parameters in PowerFactory it is calculated from the
			# calculated harmonic distortion.  This will be calculated upto and including the upper limits set in the
			# inputs for the harmonic load flow study

			# Try / except statement to allow error catching if a poor result is returned and will then be alerted
			# to user
			try:
				# res12[3:] used since at this stage the res12 format is:
				# [result type (i.e. m:HD), terminal (i.e. *.ElmTerm), H1, H2, H3, ..., Hx]
				thd = math.sqrt(sum(i*i for i in res12[3:]))

			except TypeError:
				logger.error(('Unable to calculate the THD since harmonic results retrieved from results variable {} ' +
							 ' have come out in an unexpected order and now contain a string \n' +
							 'The returned results <res12> are {}').format(self.hldf_results, res12))
				thd = 'NA'

			# #thd1 = re.split(r'[\\.]', res12[1])
			# #logger.info('thd1[11] = {}.ElmSubstat'.format(thd1[11]))
			# #thd2 = app.GetCalcRelevantObjects(thd1[11] + ".ElmSubstat")
			# #thdz = False
			# #if thd2[0] is not None:
			# #	thd3 = thd2[0].GetContents()
			# #	for thd4 in thd3:
			# #		if (thd1[13] + ".ElmTerm") in str(thd4):
			# #			logger.info('thd4 = {}'.format(thd4))
			# #			str_thd = thd4.GetAttribute('m:THD')
			# #			thdz = True
			# #elif thd2[0] is not None or thdz == False:
			# #	str_thd = "NA"
			# #res12.insert(2, str_thd)														# Insert THD
			res12.insert(2, thd)															# Insert THD
			# #res12.insert(2, New_Contingency_List[count][0])							# Op scenario
			res12.insert(2, self.cont_name)												# Op scenario
			res12.insert(2, self.sc_name)												# Study case description
			res12.pop(5)

		self.hrm_scale = hrm_scale

		return hrm_res


class PFProject:
	""" Class contains reference to a project, results folder and associated task automation file"""
	def __init__(self, name, prj, task_auto, folders):
		"""
			Initialise class
		:param str name:  project name
		:param object prj:  project reference
		:param object task_auto:  task automation reference
		:param list folders:  List of folders created as part of project, these will be deleted at end of study
		"""

		# TODO: When initialising find the initial study case, operating scenario and variations
		# TODO: So that they can be restored when the project folders are deleted

		self.name = name
		self.prj = prj
		self.task_auto = task_auto
		self.sc_cases = []
		self.folders = folders

	def process_fs_results(self):
		""" Loop through each study case cls and process results files
		:return list fs_res
		"""
		fs_res = []
		for sc_cls in self.sc_cases:
			# #sc_cls.sc.Activate()
			fs_res.extend(sc_cls.process_fs_results())
			# #sc_cls.sc.Deactivate()
		return fs_res

	def process_hrlf_results(self, logger):
		""" Loop through each study case cls and process results files
		:return list hrlf_res:
		"""
		hrlf_res = []
		for sc_cls in self.sc_cases:
			# #sc_cls.sc.Activate()
			hrlf_res.extend(sc_cls.process_hrlf_results(logger))
			# #sc_cls.sc.Deactivate()
		return hrlf_res