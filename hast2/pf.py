"""
#######################################################################################################################
###													pf																###
###		Script deals with writing data to PowerFactory and any processing that takes place which requires 			###
###		interacting with power factory																				###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###		project JI6973 for EirGrid project PSPF010 - Specialise Support in Power Quality Analysis during 2018		###
###																													###
#######################################################################################################################
"""

import math
import hast2.constants as constants
import multiprocessing
import unittest

# Meta Data
__author__ = 'David Mills'
__version__ = '1.3a'
__email__ = 'david.mills@pscconsulting.com'
__phone__ = '+44 7899 984158'
__status__ = 'In Development - Alpha'

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

def retrieve_results(elmres, res_type, write_as_df=False):			# Reads results into python lists from results file
	"""
		Reads results into python lists from results file for processing to add to Excel		
		TODO:  Adjust processing of results to write into a DataFrame for easier extraction to Excel / manipulation
	:param powerfactory.Results elmres: handle for powerfactory results file 
	:param int res_type: Type of results being dealt with	
	:param bool write_as_df:  (optional=False) If set to True will return results in a DataFrame
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

def add_filter_to_pf(_app, filter_name, filter_ref, q, freq, logger):
	"""
		Adds the filter detailed to the PF model
	:param _app: handle to power factory application
	:param excel_writing.SubstationFilter filter_ref:  Handle to SubstationFilter class form HAST import
	:param float q:  MVAR value for filter
	:param float freq:  Frequency value for filter
	:param str filter_name:  Name of filter being added which includes associated contingency
	:param logger:  Handle for logger in case of any error messages
	:return None:
	"""
	# Check that is supposed to be added
	if not filter_ref.include:
		logger.error(
			'Filter <{}> is set to not be included but has been attempted to be added, there is an error somewhere'
		.format(filter_ref.name))
		raise IOError('An error has occured trying to add a filter which should not be added')

	hdl_substation = _app.GetCalcRelevantObjects(filter_ref.sub)

	hdl_filter = create_object(location=hdl_substation[0],
						   pfclass=constants.PowerFactory.pf_filter,
						   name=filter_name)


	c = constants.PowerFactory
	# Add cubicle for filter
	hdl_terminal = hdl_substation[0].GetContents(filter_ref.term)
	hdl_cubicle = create_object(location=hdl_terminal[0],
							   pfclass=c.pf_cubicle,
							   name=filter_name)

	# Set attributes for new filter
	hdl_filter.SetAttribute(c.pf_shn_term, hdl_cubicle)
	hdl_filter.SetAttribute(c.pf_shn_voltage, filter_ref.nom_voltage)
	hdl_filter.SetAttribute(c.pf_shn_type, filter_ref.type)
	hdl_filter.SetAttribute(c.pf_shn_q, q)
	# For some reason need to set frequency and tuning order
	hdl_filter.SetAttribute(c.pf_shn_freq, freq)
	hdl_filter.SetAttribute(c.pf_shn_tuning, freq / constants.nom_freq)
	# For some reason need to set q factor
	hdl_filter.SetAttribute(c.pf_shn_qfactor, filter_ref.quality_factor)
	hdl_filter.SetAttribute(c.pf_shn_qfactor_nom, filter_ref.quality_factor * (constants.nom_freq / freq))

	hdl_filter.SetAttribute(c.pf_shn_rp, filter_ref.resistance_parallel)
	logger.debug('Filter {} added to substation {} with Q = {} MVAR and resonant frequency = {} Hz'
				 .format(filter_name, hdl_cubicle, q, freq))

	# TODO:  Rather that writing messages to confirm this could instead validate using
	logger.info(hdl_filter)
	logger.info('Connected cubicle = {} = {}'.format(hdl_cubicle, hdl_filter.GetAttribute(c.pf_shn_term)))
	logger.info('Nominal voltage = {}kV = {}kV'.format(filter_ref.nom_voltage, hdl_filter.GetAttribute(c.pf_shn_voltage)))
	logger.info('Shunt type = {} = {}'.format(filter_ref.type, hdl_filter.GetAttribute(c.pf_shn_type)))
	logger.info('Shunt Q = {}MVAR = {}MVAR'.format(q, hdl_filter.GetAttribute(c.pf_shn_q)))
	logger.info('Shunt tuning frequency = {:.2f}Hz = {:.2f}Hz'.format(freq, hdl_filter.GetAttribute(c.pf_shn_freq)))
	logger.info('Shunt tuning order = {:.1f} = {:.1f}'.format(freq/constants.nom_freq, hdl_filter.GetAttribute(c.pf_shn_tuning)))
	logger.info('Shunt Q factor = {} = {}'.format(filter_ref.quality_factor, hdl_filter.GetAttribute(c.pf_shn_qfactor)))
	logger.info('Shunt Rp = {} = {}'.format(filter_ref.resistance_parallel, hdl_filter.GetAttribute(c.pf_shn_rp)))

	# Update drawing
	_app.ExecuteCmd('grp/abi')

	return None

def set_max_processes(_app, logger):
	"""

		DOESN'T WORK - Requires PowerFactory to run in Administrator to change settings
		TODO: To fix would need to close and reopen PowerFactory as Admin, make changes
		TODO: then close and reopen with correct user

		Function will limit the number of processes to ensure that the maximum available
		RAM on the machine is not used up.

		Approach is to determine the amount of RAM that is being used up and then assume
		that if the current process is multiplied to keep some RAM available.  This is a
		pessimistic assumption but should ensure stability.

		Requires the wmi module, if not available then will limit to maximum of either:
			- 3 processes
			- no. of cores - 1 process
		<https://stackoverflow.com/questions/2017545/get-memory-usage-of-computer-in-windows-with-python>

	:param handle _app:  reference to the powerfactory application
	:param logger.LOGGER logger:  handle for the logging engine
	:return int new_slave_num:  Max processes that have been set to run, if 0 then error
	"""

	# Determine number of cores available
	max_cpu = multiprocessing.cpu_count() - constants.cpu_keep_free
	logger.debug('There are {} CPUs available for parallel processing'.format(max_cpu))

	# Obtain wmi module
	try:
		import wmi
		wmi_available = True
	except ImportError:
		logger.error('python module, wmi has not been installed and so have to limit cores'
					 'based on assumed maximum')
		wmi_available = False

	# Determine maximum available RAM
	if wmi_available:
		# Connect to computer
		comp = wmi.WMI()

		# Determine maximum physical memory in bytes
		for i in comp.Win32_ComputerSystem():
			memory_total = float(i.TotalPhysicalMemory)

		# Determine maximum free memory
		for os in comp.Win32_OperatingSystem():
			memory_free = float(os.FreePhysicalMemory)

		logger.debug('Total memory = {} bytes'.format(memory_total))
		logger.debug('Free memory = {} bytes'.format(memory_free))
		# Calculate max number of processes
		max_processes = int(memory_total/memory_free)
		logger.debug('Max processes = {}'.format(max_processes))

	else:
		# If not able to determine how much ram is available will limit the maximum number of
		# processes to a constant
		max_processes = constants.default_max_processes
		logger.debug('Not able to calculate available memory so max processes = {}'.format(max_processes))

	num_processes_to_set = max(max_cpu, max_processes)
	logger.debug('Max of processes / cores to use in PowerFactory is {}'.format(num_processes_to_set))

	# Set parallel processing restriction in powerfactory
	current_slave_num = _app.GetNumSlave()
	logger.info('Currently set to use {} slaves'.format(current_slave_num))
	_app.SetNumSlave(num_processes_to_set)
	logger.info('Set to use {} slaves'.format(num_processes_to_set))
	new_slave_num = _app.GetNumSlave()
	logger.info('Validating that has bene set to use {} slaves'.format(new_slave_num))

	# Return new_slave_number if a success
	if current_slave_num == new_slave_num:
		return new_slave_num
	else:
		return 0

class PFStudyCase:
	""" Class containing the details for each new study case created """
	def __init__(self, full_name, list_parameters, cont_name, sc, op, prj, task_auto, uid, filter_name=None,
				 base_case=False):
		"""
			Initialises the class with a list of parameters taken from the Study Settings import
		:param str full_name:  Full name of study case containing base case and contingency
		:param list list_parameters:  List of parameters from the Base_Scenarios inputs sheet
		:param str cont_name:  Name of contingency being considered
		:param str filter_name: (optional=None) Name of filter that has been included if applicable
		:param object sc:  Handle to newly created study case
		:param object op:  Handle to newly created operational scenario
		:param object prj:  Handle to project in which this study case is contained
		:param object task_auto:  Handle to the Task Automation object created for this project studies
		:param string uid:  Unique identifier time added to new files created
		:param bool base_case:  True / False on whether this is a base case study case, i.e. with no contingencies applied
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
		self.filter_name = filter_name
		self.base_case = base_case

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

		# Dictionary for looking up frequency scan results
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

	def process_fs_results(self, logger=None):
		"""
			Function extracts and processes the load flow results for this study case
		:param logger:  (optional=None) handle for logger to allow message reporting
		:return list fs_res
		"""
		c = constants.ResultsExtract

		# Insert data labels into frequency data to act as row labels for data
		fs_scale, fs_res = retrieve_results(self.fs_results, 0)
		fs_scale[0:2] = [
			c.lbl_StudyCase,
			c.lbl_Contingency,
			c.lbl_Filter_ID,
			c.lbl_FullName,
			c.lbl_Frequency]
		self.fs_scale = fs_scale

		# fs_scale.insert(1,"Frequency in Hz")										# Arranges the Frequency Scale
		# fs_scale.insert(1,"Scale")
		# fs_scale.pop(3)

		# Insert additional details for each result
		for res in fs_res:
			# Results inserted to align with labels above
			res[0:1] = [self.base_name,
						self.cont_name,
						self.filter_name,
						self.name,
						res[0]]

			# #Insert contingency / filter name (if exists)
			# #if self.filter_name is not None:
			# #	# If filter name exists then filter name is used
			# #	res.insert(1, self.filter_name)
			# #else:
			# #	# If not then contingency name is used
			# #	res.insert(1, self.cont_name)  # Contingency name

			# # Using base_name as description of study_case
			# #res.insert(1, self.base_name)  # Study case description

		logger.debug('Frequency scan results for study <{}> extracted from PowerFactory'
					 .format(self.name))

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

		# Populated with the base study case
		self.sc_base = None

	def process_fs_results(self, logger=None):
		""" Loop through each study case cls and process results files
		:return list fs_res
		"""
		fs_res = []
		for sc_cls in self.sc_cases:
			# #sc_cls.sc.Activate()
			fs_res.extend(sc_cls.process_fs_results(logger=logger))
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

	def ensure_active_study_case(self, app):
		"""
			Function checks if there is an active study case and if not will activate the study case deemed to be
			the base case to ensure that there is one active for terminal checking
		:param powerfactory.app app:  Handle to the PowerFactory application
		:return bool success:
		"""
		# Get handle for active study case from PowerFactory
		study = app.GetActiveStudyCase()

		# If no study case has been activated then activate the base case
		if study is None:
			success = not self.sc_base.sc.Activate()
		else:
			success = True

		return success


#  ----- UNIT TESTS -----
# Doesn't work because requires PowerFactory to run in Administrator mode to change settings
# class TestPowerFactoryHandling(unittest.TestCase):
# 	"""
# 		UnitTest to test the PowerFactory handling
# 	"""
# 	@classmethod
# 	def setUpClass(cls):
# 		import logging
# 		import sys
# 		import os
#
# 		# PowerFactory Python path for unittest
# 		path_pf_python = "C:\\Program Files\\DIgSILENT\\PowerFactory 2016 SP5\\Python\\3.4"
# 		path_pf = r'C:\Program Files\DIgSILENT\PowerFactory 2016 SP5'
# 		sys.path.append(path_pf_python)
# 		os.environ['PATH'] = path_pf + ';' + os.environ['PATH']
#
# 		import powerfactory
#
# 		# Initialise a logger for unittest
# 		cls.logger = logging.getLogger('UNIT TEST')
# 		cls.logger.setLevel(logging.DEBUG)
#
# 		# Get handle for powerfactory application
# 		cls.logger.debug('Establishing connection to PowerFactory, may take a while')
# 		cls.app = powerfactory.GetApplication()
# 		if cls.app is None:
# 			cls.logger.critical('Not able to access PowerFactory')
# 			assert False, 'Class not setup properly'
#
#
# 	def test_max_processes_setup(self):
# 		"""
# 			Tests that maximum number of powerfactory processes can be established
# 		"""
# 		success = set_max_processes(_app=self.app, logger=self.logger)
# 		self.assertTrue(success>0)
#
# 	@classmethod
# 	def tearDownClass(cls):
# 		# Force PowerFactory module to be released
# 		cls.app = None