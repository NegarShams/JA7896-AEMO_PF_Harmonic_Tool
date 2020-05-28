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

import os
import sys
import math
import pscharmonics.constants as constants
import pscharmonics.file_io as file_io
import multiprocessing
import time
import logging
import distutils.version
import pandas as pd

# powerfactory will be defined after initialisation by the PowerFactory class
powerfactory = None
app = None

# Meta Data
__author__ = 'David Mills'
__version__ = '2.1.2'
__email__ = 'david.mills@pscconsulting.com'
__phone__ = '+44 7899 984158'
__status__ = 'In Development - Beta'


def create_object(location, pfclass, name):  # Creates a database object in a specified location of a specified class
	"""
		Creates a database object in a specified location of a specified class
	:param powerfactory.Location location: Location in which new object should be created 
	:param str pfclass: Type of element to be created 
	:param str name: Name to be given to new object 
	:return powerfactory.Object _new_object: Handle to object returns 
	"""
	# Check if already exists before creating a new object
	existing_object = location.GetContents('{}.{}'.format(name, pfclass))
	if existing_object:
		_new_object = existing_object[0]
		already_existed = True
	else:
		_new_object = location.CreateObject(pfclass, name)
		already_existed = False
	return _new_object, already_existed

def retrieve_results(elmres, res_type, write_as_df=False):  # Reads results into python lists from results file
	"""
		Reads results into python lists from results file for processing to add to Excel
		TODO:  Adjust processing of results to write into a DataFrame for easier extraction to Excel / manipulation
	:param powerfactory.Results elmres: handle for powerfactory results file
	:param int res_type: Type of results being dealt with
	:param bool write_as_df:  (optional=False) If set to True will return results in a DataFrame
	:return (list, list), (scale, results):
	"""
	# Note both column number and row start at 1.
	# The first column is usually the scale ie timestep, frequency etc.
	# The columns are made up of Objects from left to right (ElmTerm, ElmLne)
	# The Objects then have sub variables (m:R, m:X etc)
	# TODO: This processing is slow, 20seconds per study, improve data extraction
	elmres.Load()
	cno = elmres.GetNumberOfColumns()  # Returns number of Columns
	rno = elmres.GetNumberOfRows()  # Returns number of Rows in File
	results = []
	for i in range(cno):
		column = []
		p = elmres.GetObject(i)  # Object
		d = elmres.GetVariable(i)  # Variable
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

	df = pd.DataFrame(results).transpose()
	if write_as_df:
		return df
	else:
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
	:param file_io.FilterDetails filter_ref:  Handle to FilterDetails class form HAST import
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

	hdl_filter, _ = create_object(location=hdl_substation[0],
								  pfclass=constants.PowerFactory.pf_filter,
								  name=filter_name)

	c = constants.PowerFactory
	# Add cubicle for filter
	hdl_terminal = hdl_substation[0].GetContents(filter_ref.term)
	hdl_cubicle, _ = create_object(location=hdl_terminal[0],
								   pfclass=c.pf_cubicle,
								   name=filter_name)

	# Set input mode to design mode (DES)
	hdl_filter.SetAttribute(c.pf_shn_inputmode, c.pf_shn_selectedinput)

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
	# #hdl_filter.SetAttribute(c.pf_shn_qfactor_nom, filter_ref.quality_factor * (constants.nom_freq / freq))
	hdl_filter.SetAttribute(c.pf_shn_qfactor_nom, filter_ref.quality_factor)

	hdl_filter.SetAttribute(c.pf_shn_rp, filter_ref.resistance_parallel)
	logger.debug('Filter {} added to substation {} with Q = {} MVAR and resonant frequency = {} Hz'
				 .format(filter_name, hdl_cubicle, q, freq))

	logger.info('Filter {} added to model'.format(hdl_filter))
	logger.debug('Filter input mode set to: {} and should be {}'.format(hdl_filter.GetAttribute(c.pf_shn_inputmode),
																		c.pf_shn_selectedinput))
	logger.debug('Connected cubicle = {} = {}'.format(hdl_cubicle, hdl_filter.GetAttribute(c.pf_shn_term)))
	logger.debug(
		'Nominal voltage = {}kV = {}kV'.format(filter_ref.nom_voltage, hdl_filter.GetAttribute(c.pf_shn_voltage)))
	logger.debug('Shunt type = {} = {}'.format(filter_ref.type, hdl_filter.GetAttribute(c.pf_shn_type)))
	logger.debug('Shunt Q = {}MVAR = {}MVAR'.format(q, hdl_filter.GetAttribute(c.pf_shn_q)))
	logger.debug('Shunt tuning frequency = {:.2f}Hz = {:.2f}Hz'.format(freq, hdl_filter.GetAttribute(c.pf_shn_freq)))
	logger.debug('Shunt tuning order = {:.1f} = {:.1f}'.format(freq / constants.nom_freq,
															   hdl_filter.GetAttribute(c.pf_shn_tuning)))
	logger.debug(
		'Shunt Q factor = {} = {}'.format(filter_ref.quality_factor, hdl_filter.GetAttribute(c.pf_shn_qfactor)))
	logger.debug('Shunt Rp = {} = {}'.format(filter_ref.resistance_parallel, hdl_filter.GetAttribute(c.pf_shn_rp)))

	# Update drawing so will appear if go and investigate study case
	_app.ExecuteCmd('grp/abi')

	return None

def add_vars_res(elmres, element, res_vars):	# Adds the results variables to the results file
	"""
		Adds the results variables to the results file
	:param powerfactory.DataObject elmres: Results file to be updated
	:param powerfactory.DataObject element: Element to be added
	:param tuple res_vars: 
	:return: None
	"""
	# Loop through adding each results variable to the results element
	for x in res_vars:
		elmres.AddVariable(element, x)

	return None

def check_if_object_exists(location, name):  # Check if the object exists
	# _new_object used instead of new_object to avoid shadowing
	new_object = location.GetContents(name)
	return new_object[0]

def create_mutual_name(term1, term2):
	"""
		Function creates a name for the mutual terminals ensuring does not exceed the maximum name length
		and returns the selected name alongside the planned name
	:param str term1:  Terminal 1 name
	:param str term2:  Terminal 2 name
	:return (str, str), planned_name, used_name:  The name that was planned and then what was actually used
	"""

	# Constants for terminal names
	c = constants.Terminals

	planned_name = '{}{}{}'.format(term1, c.join_char, term2)

	# Overall length determination
	if len(planned_name) > c.max_coupled_length:
		# Name is too long so need to trim characters from each terminal
		max_terminal_length = math.ceil(float(c.max_coupled_length - len(c.join_char)) / 2.0)

		term1 = term1[:max_terminal_length]
		term2 = term2[:max_terminal_length]
	used_name = '{}{}{}'.format(term1, c.join_char, term2)

	return planned_name, used_name

def create_mutual_elm(location, name, bus1, bus2):		# Creates Mutual Impedance between two terminals
	"""
		Create mutual impedance between two terminals
	:param powerfactory.DataObject location: Handle for location to save mutual impedance 
	:param str name: Name for mutual impedance 
	:param powerfactory.DataObject bus1: Terminal 1 of mutual impedance
	:param powerfactory.DataObject bus2: Terminal 2 of mutual impedance
	:return: powerfactory.DataObject  elmmut: Handle for mutual impedance
	"""
	# elmmut = app.GetFromStudyCase(name + )				# Get relevant object or create if it doesn't exist
	elmmut, _ = create_object(
		location=location,
		pfclass=constants.PowerFactory.pf_mutual,
		name=name
	)
	elmmut.loc_name = name
	elmmut.bus1 = bus1
	elmmut.bus2 = bus2
	elmmut.outserv = 0
	return elmmut


class PFStudyCase:
	""" Class containing the details for each study case contained within a project """

	def __init__(self, name, sc, op, prj, base_case=False, res_pth=str()):
		"""
			Initialises the class with a list of parameters taken from the Study Settings import
		:param str name:  Name of study case
		:param powerfactory.DataObject sc:  Handle to the study_case
		:param powerfactory.DataObject op:  Handle to the operating scenario
		:param powerfactory.DataObject prj: Handle to the project this case belongs to
		:param bool base_case: (optional=False) - Set to True for the base cases
		:param str res_pth: (optional=str()) - This is the path that the processed results will be saved in
		"""


		self.logger = logging.getLogger(constants.logger_name)

		# Status checker on whether this is the base_case study case.  If true then certain functions and additional
		# data sets are populated
		self.base_case = base_case

		# Unique name for this studycase
		self.name = name

		# Status set to False in combination or study_case and operating_scenario not convergent
		self.ldf_convergent = True

		# Reference to powerfactory handle for study case
		self.sc = sc
		self.op = op
		self.prj = prj

		self.active = False

		# Handles that will be populated with the relevant commands
		self.ldf = None
		self.fs = None
		self.fs_export_cmd = None
		self.cont_analysis = None

		# Reference to the results file that will be created by the frequency sweep
		self.fs_results = None
		self.cont_results = None

		# If no results path is provided then warn user and saved results to same folder as the script
		if not res_pth:
			res_pth = os.path.abspath(os.path.join(os.path.abspath(__file__), '..'))
		self.res_pth = res_pth

		# List of paths that contain the export files
		self.fs_result_exports = list()

		# DataFrame that will be populated with status of each contingency run, only created for the base_case as for
		# the actual contingency cases analysis is run individually on each study case / operating scenario combination
		if self.base_case:
			self.df = pd.DataFrame(columns=constants.Contingencies.df_columns)

		# self.base_name = list_parameters[0]
		# self.prj_name = list_parameters[1]
		# self.sc_name = remove_string_endings(astring=list_parameters[2], trailing='.IntCase')
		# self.op_name = remove_string_endings(astring=list_parameters[3], trailing='.IntScenario')
		# self.cont_name = cont_name
		# self.uid = uid
		# self.filter_name = filter_name
		# self.base_case = base_case
		# self.res_pth = results_pth
		#
		# # Get logger
		# self.logger = logging.getLogger(constants.logger_name)
		#
		# # Handle for study cases that will require activating
		# self.sc = sc
		# self.op = op
		# self.prj = prj
		# self.task_auto = task_auto
		#
		# # Attributes set during study completion
		# self.ldf = None
		# self.frq = None
		# self.hldf = None
		# self.frq_export_com = None
		# self.hldf_export_com = None
		# self.results = None
		# self.hldf_results = None
		# self.com_res = None
		# self.fs_scale = []
		# self.hrm_scale = []
		#
		# # Dictionary for looking up frequency scan results
		# self.fs_res = dict()
		#
		# # Paths for frequency and hlf results that are exported
		# self.fs_result_exports = []
		# self.hldf_result_exports = []

	def toggle_state(self, deactivate=False):
		"""
			Function to toggle the state of the study case and operating scenario
		:param bool deactivate: (optional=False) - Set to True to deactivate
		:return None:
		"""

		if deactivate and self.active:
			# Confirm this study case is the active study case before trying to deactivate
			active_sc = app.GetActiveStudyCase()
			if active_sc == self.sc:
				# Deactivate study case
				err = self.sc.Deactivate()
				self.active = False
			else:
				err = 0
		elif not deactivate and not self.active:
			# Activate both study case and operating scenario
			err = self.sc.Activate()
			# TODO: Confirm correct operating scenario is actually being activated
			err = self.op.Activate() + err
			self.active = True
		else:
			err = 0
			self.logger.debug('Study case {} already either deactivated / activated'.format(self.name))

		if err > 0 and deactivate:
			self.logger.error('Unable to deactivate the study case: {}'.format(self.sc))
		elif err > 0:
			self.logger.error('Unable to activate either the study case {} or operating scenario {}'.format(
				self.sc, self.op)
			)

		return None

	def create_load_flow(self, lf_settings):
		"""
			Create a load flow command in the study case so that the same settings will be run with the
			frequency scan and HAST file so that there are no issues with non-convergence.
		:param pscconsulting.file_io.LFSettings lf_settings:  Existing load flow settings
		:return None:
		"""
		# Name that is used for custom ldf command
		ldf_name = '{}_{}'.format(constants.General.cmd_lf_leader, constants.uid)

		# If input values have been provided for an existing command then copy that one
		ldf = None

		# Check if command has already been created and if has then just needs assigning
		existing_ldf = self.sc.GetContents('{}.{}'.format(ldf_name, constants.PowerFactory.ldf_command))
		if len(existing_ldf) > 0:
			ldf = existing_ldf[0]

		elif lf_settings:
			if lf_settings.cmd:
				ldf = self.sc.GetContents(lf_settings.cmd)
				# Check if command exists and if so copy that one with a new name
				if len(ldf) == 0:
					self.logger.warning(
						(
							'Not able to find load flow command {} in study case {}, provided settings will be used'
						).format(lf_settings.cmd, self.sc)
					)
				else:
					ldf = self.sc.AddCopy(ldf[0], ldf_name)

			if not ldf and not lf_settings.settings_error:
				# Populate settings based on provided inputs
				# See if load flow command already existed and if not create a new one
				ldf, _ = create_object(
					location=self.sc,
					pfclass=constants.PowerFactory.ldf_command,
					name=ldf_name)

				# Get handle for load flow command from study case
				# Basic
				ldf.iopt_net = lf_settings.iopt_net  # Calculation method (0 Balanced AC, 1 Unbalanced AC, DC)

				# Added in Automatic Tapping of PSTs but for backwards compatibility will ensure can work when less than 1
				ldf.iPST_at = lf_settings.iPST_at  # Automatic Tap Adjustment of Phase Shifters
				ldf.iopt_plim = lf_settings.iopt_plim  # Consider Active Power Limits

				# Voltage and Reactive Power Regulation
				ldf.iopt_at = lf_settings.iopt_at  # Automatic Tap Adjustment
				ldf.iopt_asht = lf_settings.iopt_asht  # Automatic Shunt Adjustment
				ldf.iopt_lim = lf_settings.iopt_lim  # Consider Reactive Power Limits
				ldf.iopt_limScale = lf_settings.iopt_limScale  # Consider Reactive Power Limits Scaling Factor

				# Temperature Dependency
				ldf.iopt_tem = lf_settings.iopt_tem  # Temperature Dependency: Line Cable Resistances
				# 													(0 ...at 20C, 1 at Maximum Operational Temperature)

				# Load Options
				ldf.iopt_pq = lf_settings.iopt_pq  # Consider Voltage Dependency of Loads
				ldf.iopt_fls = lf_settings.iopt_fls  # Feeder Load Scaling

				ldf.iopt_sim = lf_settings.iopt_sim  # Consider Coincidence of Low-Voltage Loads
				ldf.scPnight = lf_settings.scPnight  # Scaling Factor for Night Storage Heaters

				# Active Power Control
				ldf.iopt_apdist = lf_settings.iopt_apdist  # Active Power Control (0 as Dispatched, 1 According to Secondary Control,
				# 2 According to Primary Control, 3 According to Inertias)

				ldf.iPbalancing = lf_settings.iPbalancing  # (0 Ref Machine, 1 Load, Static Gen, Dist slack by loads, Dist slack by Sync,

				# Find busbar in system
				lf_settings.find_reference_terminal(app=app)
				ldf.rembar = lf_settings.rembar  # Reference machine

				ldf.phiini = lf_settings.phiini  # Angle

				# Advanced Options
				ldf.i_power = lf_settings.i_power  # Load Flow Method ( NR Current, 1 NR (Power Eqn Classic)
				ldf.iopt_notopo = lf_settings.iopt_notopo  # No Topology Rebuild
				ldf.iopt_noinit = lf_settings.iopt_noinit  # No initialisation
				ldf.utr_init = lf_settings.utr_init  # Consideration of transformer winding ratio
				ldf.maxPhaseShift = lf_settings.maxPhaseShift  # Max Transformer Phase Shift
				ldf.itapopt = lf_settings.itapopt  # Tap Adjustment ( 0 Direct, 1 Step)
				ldf.krelax = lf_settings.krelax  # Min Controller Relaxation Factor

				ldf.iopt_stamode = lf_settings.iopt_stamode  # Station Controller (0 Standard, 1 Gen HV, 2 Gen LV
				ldf.iopt_igntow = lf_settings.iopt_igntow  # Modelling Method of Towers (0 With In/ Output signals, 1 ignore couplings, 2 equation in lines)
				ldf.initOPF = lf_settings.initOPF  # Use this load flow for initialisation of OPF
				ldf.zoneScale = lf_settings.zoneScale  # Zone Scaling ( 0 Consider all loads, 1 Consider adjustable loads only)

				# Iteration Control
				ldf.itrlx = lf_settings.itrlx  # Max No Iterations for Newton-Raphson Iteration
				ldf.ictrlx = lf_settings.ictrlx  # Max No Iterations for Outer Loop
				ldf.nsteps = lf_settings.nsteps  # Max No Iterations for Number of steps

				ldf.errlf = lf_settings.errlf  # Max Acceptable Load Flow Error for Nodes
				ldf.erreq = lf_settings.erreq  # Max Acceptable Load Flow Error for Model Equations
				ldf.iStepAdapt = lf_settings.iStepAdapt  # Iteration Step Size ( 0 Automatic, 1 Fixed Relaxation)
				ldf.relax = lf_settings.relax  # If Fixed Relaxation factor
				ldf.iopt_lev = lf_settings.iopt_lev  # Automatic Model Adaptation for Convergence

				# Outputs
				ldf.iShowOutLoopMsg = lf_settings.iShowOutLoopMsg  # Show 'outer Loop' Messages
				ldf.iopt_show = lf_settings.iopt_show  # Show Convergence Progress Report
				ldf.num_conv = lf_settings.num_conv  # Number of reported buses/models per iteration
				ldf.iopt_check = lf_settings.iopt_check  # Show verification report
				ldf.loadmax = lf_settings.loadmax  # Max Loading of Edge Element
				ldf.vlmin = lf_settings.vlmin  # Lower Limit of Allowed Voltage
				ldf.vlmax = lf_settings.vlmax  # Upper Limit of Allowed Voltage
				# ldf.outcmd =  load_flow_settings[42-offset]          		# Output
				ldf.iopt_chctr = lf_settings.iopt_chctr  # Check Control Conditions
				# ldf.chkcmd = load_flow_settings[44-offset]            	# Command

				# Load Generation Scaling
				ldf.scLoadFac = lf_settings.scLoadFac  # Load Scaling Factor
				ldf.scGenFac = lf_settings.scGenFac  # Generation Scaling Factor
				ldf.scMotFac = lf_settings.scMotFac  # Motor Scaling Factor

				# Low Voltage Analysis
				ldf.Sfix = lf_settings.Sfix  # Fixed Load kVA
				ldf.cosfix = lf_settings.cosfix  # Power Factor of Fixed Load
				ldf.Svar = lf_settings.Svar  # Max Power Per Customer kVA
				ldf.cosvar = lf_settings.cosvar  # Power Factor of Variable Part
				ldf.ginf = lf_settings.ginf  # Coincidence Factor
				ldf.i_volt = lf_settings.i_volt  # Voltage Drop Analysis (0 Stochastic Evaluation,
				#														 						1 Maximum Current Estimation)

				# Advanced Simulation Options
				ldf.iopt_prot = lf_settings.iopt_prot  # Consider Protection Devices ( 0 None, 1 all, 2 Main, 3 Backup)
				ldf.ign_comp = lf_settings.ign_comp  # Ignore Composite Elements

				self.logger.debug(
					(
						'Load flow settings for study case <{}> based on settings in inputs spreadsheet and '
						'detailed in load flow command <{}>'
					).format(self.sc, ldf)
					)

		# If ldf still hasn't been defined then use default load flow
		if not ldf:
			# Get default load flow command, copy and rename
			def_ldf = self.sc.GetContents('*.{}'.format(constants.PowerFactory.ldf_command))[0]
			ldf = self.sc.AddCopy(def_ldf, ldf_name)
			self.logger.warning(
				(
					'Not able to use provided load flow settings or existing load flow command for study case {} and '
					'therefore a new command <{}> has been created based on the default command <{}>'
				).format(self.sc, def_ldf, ldf)
			)

		self.ldf = ldf

		self.delete_sc_objects(pf_cmd=self.ldf, pf_type=constants.PowerFactory.ldf_command)

		return None

	def create_freq_sweep(self, fs_settings):
		"""
			Create a frequencys weep command in the study case so that the same settings will be run for all
			subsequent study cases.
		:param pscconsulting.file_io.FSSettings fs_settings:  Settings to use
		:return None:
		"""
		# Name that is used for custom ldf command
		fs_name = '{}_{}'.format(constants.General.cmd_fs_leader, constants.uid)

		# Confirm load flow and results file has already been defined since needed for output settings
		if self.ldf is None:
			self.logger.error(
				(
					'Not possible to create frequency scan for study case {} since no load flow settings have yet'
					'been determined.  This could be a scripting issue or an error finding a suitable load flow.'
				).format(self.sc)
			)
			self.fs = None
		else:

			# If input values have been provided for an existing command then copy that one
			fs = None
			# Check if command has already been created and if has then just needs assigning
			existing_fs = self.sc.GetContents('{}.{}'.format(fs_name, constants.PowerFactory.fs_command))
			if len(existing_fs) > 0:
				fs = existing_fs[0]

			if fs_settings:

				if fs_settings.cmd:
					# If no existing command then create a new one
					fs = self.sc.GetContents(fs_settings.cmd)
					# Check if command exists and if so copy that one with a new name
					if len(fs) == 0:
						self.logger.warning(
							(
								'Not able to find frequency sweep command {} in study case {}, provided settings will be '
								'used instead.'
							).format(fs_settings.cmd, self.sc)
						)
					else:
						fs = self.sc.AddCopy(fs[0], fs_name)

				if not fs and not fs_settings.settings_error:
					# Populate settings based on provided inputs
					# See if frequency sweep command already existed and if not create a new one
					fs, _ = create_object(
						location=self.sc,
						pfclass=constants.PowerFactory.fs_command,
						name=fs_name)

					# Get handle for frequency sweep command from study case
					fs.iopt_net = fs_settings.iopt_net  # Network Representation (0=Balanced 1=Unbalanced)
					fs.fstart = fs_settings.fstart  # Impedance Calculation Start frequency
					fs.fstop = fs_settings.fstop  # Stop Frequency
					fs.fstep = fs_settings.fstep  # Step Size
					fs.i_adapt = fs_settings.i_adapt  # Automatic Step Size Adaption
					fs.frnom = fs_settings.frnom  # Nominal Frequency
					fs.fshow = fs_settings.fstop # Fixzed to be the same as the stop frequency
					fs.ifshow = float(fs_settings.fstop) / float(fs_settings.frnom)  # Harmonic Order

					# Advanced
					fs.errmax = fs_settings.errmax  # Setting for Step Size Adaption    Maximum Prediction Error
					fs.errinc = fs_settings.errinc  # Minimum Prediction Error
					fs.ninc = fs_settings.ninc  # Step Size Increase Delay
					fs.ioutall = fs_settings.ioutall  # Fixed to not include output for R, X at all nodes

					self.logger.debug(
						(
							'Load flow settings for study case <{}> based on settings in inputs spreadsheet and '
							'detailed in frequency sweep command <{}>'
						).format(self.sc, fs)
					)

			# If ldf still hasn't been defined then use default load flow
			if not fs:
				# Get default load flow command, copy and rename
				def_fs = self.sc.GetContents('*.{}'.format(constants.PowerFactory.fs_command))[0]
				fs = self.sc.AddCopy(def_fs, fs_name)
				self.logger.warning(
					(
						'Not able to use provided frequency sweep settings or existing frequency sweep command for study '
						'case {} and therefore a new command <{}> has been created based on the default command <{}>'
					).format(self.sc, def_fs, fs)
				)

			# Check if results file has already been defined otherwise define a new one
			if not self.fs_results:
				self.create_results_files()
			# Reference to results file where frequency scan results will be saved
			# fs.SetAttribute('p_resvar', self.results)
			fs.p_resvar = self.fs_results  # Results Variable

			# Frequency sweep will use the load flow command created for this study case
			fs.c_butldf = self.ldf

			# Delete all other frequency scan objects
			self.delete_sc_objects(pf_cmd=fs, pf_type=constants.PowerFactory.fs_command)

			self.fs = fs
		return None

	def pre_case_check(self):
		"""
			Function to create all the necessary contingency cases and run a pre-case check to confirm that all
			cases are convergent.  The status of each case is then updated in the DataFrame which will be exported
			at the end of the study and used as the basis for whether frequency scans should be run
		:return None:
		"""

		# Since creating contingency analysis is to confirm that model is convergent for every contingency initially
		# run a load flow study to confirm the intact condition is convergent.
		self.run_load_flow()
		c = constants.Contingencies

		if self.ldf_convergent:
			# Update dataframe to show intact system is convergent
			self.df.loc[c.intact, c.status] = True
		else:
			self.logger.error(
				(
					'The base case for study case <{}> with operating scenario <{}> is not-convergent and therefore no '
					'studies can be run.  Please check the initial case'
				).format(self.sc, self.op)
			)
			return None

	def create_cont_analysis(self, fault_cases=None, cmd=str()):
		"""
			Creates a contingency analysis command in the study case so can iterate through all contingencies
			and identify those which are not convergent.
		:param dict fault_cases:  Dictionary of fault_cases as returned by PowerFactory.create_fault_cases
		:param str cmd:  Name of command if already provided as part of input data
		:return None:
		"""
		c = constants.Contingencies

		# Name that is used for custom ldf command
		name = '{}_{}'.format(constants.General.cmd_cont_leader, constants.uid)
		cont_analysis = None

		# Confirm load flow and results file has already been defined since needed for output settings
		if self.ldf is None:
			self.logger.error(
				(
					'Not possible to create contingency analysis for study case {} since no load flow settings have yet '
					'been determined.  This could be a scripting issue or an error finding a suitable load flow.'
				).format(self.sc)
			)
			cont_analysis = None

		else:
			if cmd:
				# Get existing command from StudyCase and base analysis on that
				existing = self.sc.GetContents('{}.{}'.format(cmd, constants.PowerFactory.pf_cont_analysis))

				if len(existing) > 0:
					# Check if already exists and if so duplicate existing one so settings can be changed
					cont_analysis = self.sc.AddCopy(existing[0], name)

					# Loop through all contingencies within this defined contingency set and update the dataframe with
					# relevant details.
					outages = cont_analysis.GetContents('*.{}'.format(constants.PowerFactory.pf_outage))
					if len(outages) == 0:
						self.logger.warning(
							(
								'No outages have been defined in the provided contingency set {} as part of study case '
								'<{}>.  If individual outage details have been provided as part of the input these will '
								'be used instead and otherwise the script will fail.'
							).format(cmd, self.sc)
						)
					else:
						# Loop through each outage and update DataFrame with some relevant details
						for outage in outages:
							cont_name = outage.loc_name
							self.df.loc[cont_name, c.cont] = cont_name
							self.df.loc[cont_name, c.idx] = outage.number

				else:
					# If a command has been provided but cannot be found then display a warning message
					self.logger.warning(
						(
							'The provided contingency set {} does not exist in the study case <{}> and therefore cannot '
							'be used for contingency analysis.  If individual outage details have been provided they will'
							'be used instead otherwise the script will fail.'
						).format(cmd, self.sc)
					)
					cont_analysis = None

			if fault_cases and not cont_analysis:
				# Create Contingency Analysis command and add fault cases to it if has not been possible to establish
				# cont_analysis from the input provided by the user.
				cont_analysis, _ = create_object(
					location=self.sc,
					pfclass=constants.PowerFactory.pf_cont_analysis,
					name=name
				)

				# Loop through each fault case and create a contingency with each contingency being added to the
				# study case specific dataframe.  The status of each contingency is then updated once the initial
				# pre-case check is carried out.
				counter = 1
				for cont_name, fault in fault_cases.items():
					outage, _ = create_object(
						location=cont_analysis,
						pfclass=constants.PowerFactory.pf_outage,
						name=cont_name
					)
					# Set Outage up to represent this fault case
					outage.cpCase = fault

					# Update DataFrame with details of this contingency
					outage.number = counter
					self.df.loc[cont_name, c.idx] = outage.number
					self.df.loc[cont_name, c.cont] = cont_name
					counter += 1

			if not cont_analysis:
				# No command or fault cases provided so raise error to user
				self.logger.critical(
					(
						'No fault cases provided as input and no existing command existed in study case <{}>.'
						' The following inputs were provided:\n\t Fault Cases = {}\n\tCmd = {}'
					).format(self.sc, fault_cases, cmd)
				)
				raise SyntaxError('Incorrect inputs')

			# Set default parameters for contingency analysis to ensure aligns with load flow run
			# Run based on normal AC load flow with previously created load flow settings
			cont_analysis.iopt_Linear = 0
			cont_analysis.ldf = self.ldf
			# Ensure results are stored in the results variable
			cont_analysis.p_recnt = self.cont_results

			# If a large number of contingencies then allow parallel running of cases
			cont_analysis.iEnableParal = 1
			cont_analysis.minCntcyAC = c.parallel_threshold

			# Delete all other contingency analysis objects
			self.delete_sc_objects(pf_cmd=cont_analysis, pf_type=constants.PowerFactory.pf_cont_analysis)

		# Update DataFrame to ensure study case and operating scenario columns match this data
		self.df[c.prj] = self.prj.loc_name
		self.df[c.sc] = self.sc.loc_name
		self.df[c.op] = self.op.loc_name

		self.cont_analysis = cont_analysis
		return None

	def delete_sc_objects(self, pf_cmd, pf_type):
		"""
			Function to delete all of the other items of a particular type from the study case except the one provided
			as a reference as an input
		:param (PowerFactory.DataObject, ) pf_cmd: Reference to object(s) to be kept which are input as either a tuple
													or single item and then converted to a tuple
		:param str pf_type:  Extension of variable type to be deleted
		:return int num_deleted:  Number of objects deleted
		"""

		# Get all values of a particular type from study case
		pf_objects = self.sc.GetContents('*.{}'.format(pf_type))
		original_number = len(pf_objects)

		# Check if input is a tuple so can be iterated through
		if type(pf_cmd) is not tuple:
			pf_cmd = (pf_cmd, )

		if original_number == 0:
			# If none exist then warn user
			self.logger.warning(
				(
					'Attempted to delete all objects of type {} from study case {} except the object <{}> but none '
					'existed'
				).format(pf_type, self.sc, pf_cmd)
			)

		deleted_objects=list()
		for obj in pf_objects:
			if obj not in pf_cmd:
				obj.Delete()
				deleted_objects.append(str(obj))

		if len(deleted_objects) != original_number-len(pf_cmd):
			self.logger.warning(
				(
					'Have attempted to delete objects of type {} from the study case {} which originally consisted of \n\t'
					'{}\n'
					'But have only been able to delete {} which consisted of the following \n\t'
					'{}\n'
				).format(pf_type, self.sc, '\n\t'.join(
					[str(x) for x in pf_objects]
				), len(deleted_objects), '\n\t'.join(
					[str(x) for x in deleted_objects]))
			)
		else:
			self.logger.debug(
				(
					'Successfully deleted {} objects of type {} from study case {} except the relevant item {}'
				).format(len(deleted_objects), pf_type, self.sc, pf_cmd)
			)

		# Return an index showing the number of objects deleted
		return len(deleted_objects)

	# def process_fs_results(self, logger=None):
	# 	"""
	# 		Function extracts and processes the load flow results for this study case
	# 	:param logger:  (optional=None) handle for logger to allow message reporting
	# 	:return list fs_res
	# 	"""
	# 	c = constants.Results
	#
	# 	# Insert data labels into frequency data to act as row labels for data
	# 	fs_scale, fs_res = retrieve_results(self.fs_results, 0)
	# 	fs_scale[0:2] = [
	# 		c.lbl_StudyCase,
	# 		c.lbl_Contingency,
	# 		c.lbl_Filter_ID,
	# 		c.lbl_FullName,
	# 		c.lbl_Frequency]
	# 	self.fs_scale = fs_scale
	#
	# 	# fs_scale.insert(1,"Frequency in Hz")										# Arranges the Frequency Scale
	# 	# fs_scale.insert(1,"Scale")
	# 	# fs_scale.pop(3)
	#
	# 	# Insert additional details for each result
	# 	for res in fs_res:
	# 		# Results inserted to align with labels above
	# 		res[0:1] = [self.base_name,
	# 					self.cont_name,
	# 					self.filter_name,
	# 					self.name,
	# 					res[0]]
	#
	# 	logger.debug('Frequency scan results for study <{}> extracted from PowerFactory'
	# 				 .format(self.name))
	#
	# 	return fs_res

	def process_cont_results(self):
		"""
			Function will process the contingencies to check the results and determine which were convergent
			with the status being updated in the DataFrame.
		:return:
		"""
		c = constants.Contingencies
		df = retrieve_results(elmres=self.cont_results, res_type=0, write_as_df=True)

		# If an empty DataFrame is returned then means all contingencies failed so set status to False
		if df.empty:
			self.df.loc[:, c.status] = False
		else:
			# Set columns to be based on first index
			df.columns = df.loc[0, :]

			# Drop non-relevant rows
			df.drop(labels=[0, 1], axis=0, inplace=True)
			# Drop last row which also isn't needed
			df.drop(df.tail(1).index, inplace=True)

			# Set the index for the DataFrame based on the object number
			df.set_index(c.col_number, inplace=True)

			# Get list of contingencies which were not convergent
			for cont_number, series in df.iterrows():
				# Populate the status of the convergence
				self.df.loc[self.df[c.idx]==cont_number, c.status] = not bool(series[c.col_nonconvergent])

			self.logger.debug(
				'Processing contingency analysis results for case {}, consisting of sc {} and op {}'.format(
					self.name, self.sc, self.op
				)
			)

		return None


	# def process_hrlf_results(self, logger):
	# 	"""
	# 		Process the hrlf results ready for inclusion into spreadsheet
	# 	:return hrm_res
	# 	"""
	# 	hrm_scale, hrm_res = retrieve_results(self.hldf_results, 1)
	#
	# 	hrm_scale.insert(1, "THD")  # Inserts the THD
	# 	hrm_scale.insert(1, "Harmonic")  # Arranges the Harmonic Scale
	# 	hrm_scale.insert(1, "Scale")
	# 	hrm_scale.pop(4)  # Takes out the 50 Hz
	# 	hrm_scale.pop(4)
	# 	for res12 in hrm_res:
	# 		# Rather than retrieving THD from the calculated parameters in PowerFactory it is calculated from the
	# 		# calculated harmonic distortion.  This will be calculated upto and including the upper limits set in the
	# 		# inputs for the harmonic load flow study
	#
	# 		# Try / except statement to allow error catching if a poor result is returned and will then be alerted
	# 		# to user
	# 		try:
	# 			# res12[3:] used since at this stage the res12 format is:
	# 			# [result type (i.e. m:HD), terminal (i.e. *.ElmTerm), H1, H2, H3, ..., Hx]
	# 			thd = math.sqrt(sum(i * i for i in res12[3:]))
	#
	# 		except TypeError:
	# 			logger.error(('Unable to calculate the THD since harmonic results retrieved from results variable {} ' +
	# 						  ' have come out in an unexpected order and now contain a string \n' +
	# 						  'The returned results <res12> are {}').format(self.hldf_results, res12))
	# 			thd = 'NA'
	#
	# 		res12.insert(2, thd)  # Insert THD
	# 		res12.insert(2, self.cont_name)  # Op scenario
	# 		res12.insert(2, self.sc_name)  # Study case description
	# 		res12.pop(5)
	#
	# 	self.hrm_scale = hrm_scale
	#
	# 	return hrm_res

	def create_results_files(self):
		"""
			Function creates a results file if it does not already exist
		:return None:
		"""
		# Update FS results file
		self.fs_results, _ = create_object(
			location=self.sc,
			pfclass=constants.PowerFactory.pf_results,
			name='{}{}'.format(constants.General.cmd_fsres_leader, constants.PowerFactory.default_fs_extension)
		)

		# Update Contingency analysis results file
		self.cont_results, _ = create_object(
			location=self.sc,
			pfclass=constants.PowerFactory.pf_results,
			name='{}{}'.format(constants.General.cmd_contres_leader, constants.PowerFactory.default_fs_extension)
		)
		# Set as default results for Freq.Sweep
		self.fs_results.calTp = constants.PowerFactory.def_results_fs
		# Set as default results for contingency analysis based on AC Load Flow
		self.cont_results.calTp = constants.PowerFactory.def_results_cont
		self.cont_results.calTpSub = 0

		self.delete_sc_objects(pf_cmd=(self.fs_results, self.cont_results), pf_type=constants.PowerFactory.pf_results)
		return None

	def add_variables(self, study_settings, terminals, mutuals):
		"""
			Function adds the required variables to the fs results file based on the study settings
		:param file_io.StudySettings study_settings:  Input settings to determine which sort of results to export
		:param dict terminals:  Dictionary with references to the terminals to be included
		:param dict mutuals:  Dictionary with references to the mutuals to be included
		:return None:
		"""

		# Confirm results variable exits and if not create
		if not self.fs_results:
			self.create_studies()

		# Determine types of variables to be declaring
		c = constants.PowerFactory
		if study_settings.export_rx:
			self_variables = (c.pf_z1, c.pf_r1, c.pf_x1)

		else:
			self_variables = (c.pf_z1, )
		self.logger.debug('Self impedance results declared for: {}'.format(' - '.join(self_variables)))

		# Mutual variables to export
		mutual_variables = tuple()
		if study_settings.export_mutual:
			if study_settings.export_rx:
				mutual_variables = (c.pf_z12, c.pf_r12, c.pf_x12)

			else:
				mutual_variables = (c.pf_z12, )
			self.logger.debug('Mutual impedance results declared for: {}'.format(' - '.join(mutual_variables)))
		else:
			self.logger.debug('No mutual impedance results to be calculated')

		# TODO: At this point want to clear all existing variables from results file

		# Loop through all terminals and add
		for term_name, term in terminals.items():
			add_vars_res(
				elmres=self.fs_results,
				element=term.pf_handle,
				res_vars=self_variables
			)
			self.logger.debug(
				(
					'Terminal Named {}, relating to terminal {} added to results file {}'
				).format(term_name, term.pf_handle, self.fs_results)
			)

		if study_settings.export_mutual:
			# Add mutual impedance variables if they have been declared
			for term_name, term_handle in mutuals.items():
				add_vars_res(
					elmres=self.fs_results,
					element=term_handle,
					res_vars=mutual_variables
				)
				self.logger.debug(
					(
						'Mutual Named {}, relating to terminal {} added to results file {}'
					).format(term_name, term_handle, self.fs_results)
				)

		return None


	def create_studies(self, lf_settings=None, fs_settings=None):
		"""
			Function to either create a new command or change the reference of an existing command to results file
			associated with this study
		:param file_io.LFSettings lf_settings: (optional=None) Settings to use for the frequency sweep settings
		:param file_io.FSSettings fs_settings:  (optional=None) Settings to use for the frequency sweep settings
		:return None:
		"""
		self.create_load_flow(lf_settings=lf_settings)
		self.create_results_files()
		self.create_freq_sweep(fs_settings=fs_settings)
		self.fs_export_cmd, export_pth = self.set_results_export(result=self.fs_results)

		self.fs_result_exports.append(export_pth)
		self.logger.debug(
			(
				'For study case {}, load flow command {}, frequency scan command {} and results export {} have been '
				'created with results being exported to: {}'
			).format(self.sc, self.ldf, self.fs, self.fs_export_cmd, export_pth)
		)

	def set_results_export(self, result):
		"""
			Function will create a results export command (.ComRes) to then use to deal with exporting all the results
			into a CSV file as soon as they are completed.  They can then be processed by a different script.
		:param powerfactory.DataObject result:  handle to powerfactory result that should be extracted
		:return (powerfactory.DataObject, res_export_pth):  Handle to PF ComRes function, Full path to exported result
		"""
		res_export_path = os.path.join(self.res_pth, '{}.csv'.format(self.name))

		c = constants.PowerFactory.ComRes
		# Create com_res file to deal with extracting the results
		h_comres, _ = create_object(location=self.sc, pfclass=c.pf_comres, name=self.name)

		# Set relevant result
		h_comres.SetAttribute(c.result, result)

		# Set type as CSV and define results file
		h_comres.SetAttribute(c.export_type, 6)
		h_comres.SetAttribute(c.file, os.path.join(self.res_pth, res_export_path))
		h_comres.SetAttribute(c.separators, 1)
		h_comres.SetAttribute(c.object_head_only, 0)

		# Export values (0 = values, 1 = variable descriptors only)
		h_comres.SetAttribute(c.export_values, 0)

		# Export all variables (0 = all variables, 1 = list of variables)
		h_comres.SetAttribute(c.variables_all, 0)

		# Set time steps
		h_comres.SetAttribute(c.user_interval, 0)
		h_comres.SetAttribute(c.filtered_time, 0)
		h_comres.SetAttribute(c.shift_time, 0)

		# Set data to include
		h_comres.SetAttribute(c.element, 3)
		h_comres.SetAttribute(c.variable, 1)

		return h_comres, res_export_path

	def run_load_flow(self):
		""" Function to run the embedded load flow command
		:return bool success: Returns True / False on whether load flow was a success
		"""
		# Execute load flow, track run time and confirm success
		t1 = time.time()
		error_code = self.ldf.Execute()
		t2 = time.time() - t1
		if error_code == 0:
			self.logger.debug('\t - Load Flow calculation {} successful for {}, time taken: {:.2f} seconds'
						.format(self.ldf, self.name, t2))
			success = True
		elif error_code == 1:
			self.logger.error(('Load Flow calculation {} for {} failed due to divergence of inner loops, '
						  'time taken: {:.2f} seconds')
						 .format(self.ldf, self.name, t2))
			success = False
		elif error_code == 2:
			self.logger.error(('Load Flow calculation {} failed for {} due to divergence of outer loops, '
						  'time taken: {:.2f} seconds')
						 .format(self.ldf, self.name, t2))
			success = False
		else:
			success = False

		# Set overall status
		self.ldf_convergent = success

		return success

	def create_cases(self, sc_folder, op_folder, res_pth=str()):
		"""
			Function will loop through valid contingencies and create a new case setup to reflect that contingency
			and that will then be stored in the temporary sc and op folders
		:param powerfactory.DataObject sc_folder:  Reference to the folder to store temporary study cases in
		:param powerfactory.DataObject op_folder:  Reference to the folder to store temporary oeprating scenarios in
		:param str res_pth:  Path where all the results will be saved when the automatic study cases are run
		:return list new_cases:  List of references to the newly created study cases
		"""
		# Confirm case is deactivated
		self.toggle_state(deactivate=True)

		# If no results path is provided then use default
		if not res_pth:
			res_pth = self.res_pth

		# Loop through all contingencies in this case which are convergent
		new_cases = list()
		for cont_name, cont_case in self.df[self.df[constants.Contingencies.status]==True].iterrows():
			# Create name for new case as combination of provided name and contingency
			new_name = '{}-{}'.format(self.name, cont_name)

			# Copy the current study_case and operating scenario
			new_sc = sc_folder.AddCopy(self.sc, new_name)
			new_op = op_folder.AddCopy(self.op, new_name)

			# Create new PFStudyCase instance
			case = PFStudyCase(
				name=new_name,
				sc=new_sc,
				op=new_op,
				prj=self.prj,
				res_pth=res_pth
			)

			# Get the contingency specific to this case and apply the outage which is a PowerFactory Cont Outage element
			cont_outage = self.cont_analysis.GetContents('{}.{}'.format(cont_name, constants.PowerFactory.pf_outage))

			if len(cont_outage) == 0:
				self.logger.error(
					(
						'Unable to find outage {} in contingency analysis {} and therefore study_case {} cannot be run'
					).format(cont_name, self.cont_analysis, new_name)
				)
				continue
			else:
				# Apply the outage detailed in this Cont Outage element
				case.apply_outage(cont_outage[0])

			# Update load flow and frequency sweep commands to reflect relevant locations
			case.create_studies()

			self.logger.debug(
				(
					'New case {} created for Study Case {}, Operating Scenario {} with Contingency {}'
				).format(new_name, new_sc, new_op, cont_name)
			)
			new_cases.append(case)

		return new_cases

	def apply_outage(self, cont_outage):
		"""
			Function will carry out the switching operations detailed in the cont_outage and set all interrupted
			components to out of service.  Also ensures the relevant study case and outage scenario are
			associated together
		:param powerfactory.DataObject cont_outage:  Handle to the powerfactory DataObject
		:return None:
		"""
		# Active study case and operating scenario (also ensures these are combined together for a future activation)
		self.toggle_state()

		# Get details of all switches considered as part of contingency to Open
		for switch in cont_outage.Couplers:
			# Open switch
			switch.on_off = False

		# Close all couplers that should be closed
		for switch in cont_outage.CouplersClose:
			# Open switch
			switch.on_off = False

		# Set all interrupted elements to out of service
		# TODO: Confirm if this is necessary
		for element in cont_outage.Elms:
			element.outserv = True

		# Save operating scenario so that it is remembered in this state and if errors then raise error to user
		err = self.op.Save()
		if err == 1:
			self.logger.error(
				(
					'Unable to save the operating scenario {} after applying the outage {}, this means the results for '
					'the contingency named {} as part of study case {} will not produce reliable results'
				).format(self.op, cont_outage, cont_outage.loc_name, self.name)
			)
		else:
			self.logger.debug(
				'Successfully applied outage {} for study case {}'.format(
					cont_outage, self.sc
				)
			)

		return None


class PFProject:
	""" Class contains reference to a project, results folder and associated task automation file"""

	def __init__(self, name, df_studycases, uid, lf_settings=None, fs_settings=None, res_pth=str()):
		"""
			Initialise class
		:param str name:  project name
		:param pd.DataFrame df_studycases:  DataFrame containing all the base study cases associates with this project
		:param psconsulting.file_io.LFSettings lf_settings:  (optional=None) - If provided then these settings will be
															used and if not then default Load Flow command will be used
		:param psconsulting.file_io.FSSettings fs_settings:  (optional=None) - If provided then these settings will be
															used and if not then default Frequency Sweep command will be used
		:param pd.DataFrame df_studycases:  DataFrame containing all the base study cases associates with this project
		:param str uid:  Unique identifier given for this study
		:param str res_pth: (optional=str()) - This is the path that the processed results will be saved in
		"""
		self.logger = logging.getLogger(constants.logger_name)
		self.logger.debug('New instance for project {} being initialised'.format(name))
		# self.prj_active = False

		self.name = name
		self.uid = uid
		self.pf = PowerFactory()

		# Store details of settings
		self.lf_settings = lf_settings
		self.fs_settings = fs_settings

		# DataFrame of study cases which is populated with status for base_case
		self.df_sc = df_studycases
		# DataFrame of study cases which is populated with details of convergence status and used to determine which
		# ones to create contingencies for
		self.df_pre_case = pd.DataFrame()

		# Dictionary of terminals that have been found in this project
		self.terminals = dict()
		# Dictionary of mutual elements
		self.mutuals = dict()

		# Activate project to get power_factory instance
		self.prj = self.pf.activate_project(project_name=name)
		self.prj_active = True

		# Path where all result exports will be saved to
		self.res_pth =res_pth

		if self.prj is None:
			self.logger.warning(
				(
					'Not possible to activate project named "{}" and therefore no studies will be carried out for '
					'study cases associated with this project'
				).format(self.name)
			)
			self.exists = False
		else:
			self.exists = True

		# Get reference to project study case and operational scenario folders
		self.base_sc_folder = app.GetProjectFolder(constants.PowerFactory.pf_sc_folder_type)
		self.base_os_folder = app.GetProjectFolder(constants.PowerFactory.pf_os_folder_type)
		# self.base_var_folder = app.GetProjectFolder('scheme')

		# Get handle for all network data
		self.net_data = app.GetProjectFolder(constants.PowerFactory.pf_netdata_folder_type)
		# Get all folders which contain network elements
		self.net_data_items = self.net_data.GetContents('*.{}'.format(constants.PowerFactory.pf_network_elements))

		# Create temporary folders
		self.temp_folders = list()
		c = constants.PowerFactory
		self.sc_folder = self.create_folder(
			name='{}_{}'.format(c.temp_sc_folder, self.uid),
			location=self.base_sc_folder,
			temp=True
		)
		self.op_folder = self.create_folder(
			name='{}_{}'.format(c.temp_os_folder, self.uid),
			location=self.base_os_folder,
			temp=True
		)
		# Folder to contain the fault cases which is only created if needed
		self.fault_case_folder = None

		# Populated with the fault cases which are created from the contingencies input where necessary.
		# Only populated where a contingency analysis command is not provided instead.
		self.fault_cases = dict()

		# List is populates with handles for all of the cases run as part of the frequency scan studies
		self.cases_to_run = list()

		# Create the command for the auto tasks associated with this project
		self.task_auto = self.create_task_auto()

		# Initialise study_cases
		self.base_sc = self.initialise_study_cases()

		#
		#
		# self.task_auto = task_auto
		# self.sc_cases = []
		# self.folders = folders
		#
		# # Populated with the base study case
		# self.sc_base = None
		#
		# # If Mutual impedance data required then added here
		# self.include_mutual = include_mutual
		# self.mutual_impedance_folder = None
		# # list of mutual impedance elements in the format:
		# # [(HAST_input_name,
		# # 	mutual_impedance_name (i.e. 'from_to'),
		# # 	reference to mutual element in pf,
		# # 	reference to terminal 1 in pf,
		# # 	reference to terminal 2 in pf)
		# # ]
		# self.list_of_mutual = []
		# # List of names for which mutual impedance elements have been created in the form
		# #	[from1_to1, to1_from1, from2_to2, to2_from2, ...]
		# self.list_of_mutual_names = []
		#
		# # Network elements folder
		# self.folder_network_elements = None
		#
		# # List of terminals for results
		# self.terminals_index = None

	def initialise_study_cases(self):
		"""
			Function loops through all study_case and operational scenario combinations and creates
			duplicates that are stored in the temporary folders
		:return dict base_study_cases: Returns a list of all the base study cases that have been created and
										can be activated
		"""
		base_study_cases = dict()

		# Loop through each of the provided study cases and create a reference to the study case to be run
		for idx, df_sc in self.df_sc.iterrows():
			# Get the studycase references and ensure correct IntCase or IntScenario reference
			name = df_sc.loc[constants.StudySettings.name]
			sc_name = df_sc.loc[constants.StudySettings.studycase].replace(
				'.{}'.format(constants.PowerFactory.pf_case), '')
			os_name = df_sc.loc[constants.StudySettings.scenario].replace(
				'.{}'.format(constants.PowerFactory.pf_scenario), '')
			sc_name = '{}.{}'.format(sc_name, constants.PowerFactory.pf_case)
			os_name = '{}.{}'.format(os_name, constants.PowerFactory.pf_scenario)

			# Find handle in powerfactory for study_case
			pf_sc = self.base_sc_folder.GetContents(sc_name)
			if len(pf_sc) == 0:
				# Study case doesn't exist so alert user and skip to next
				self.logger.error(
					(
						'Study Case {} cannot be found in PowerFactory folder {}, no studies will be carried out '
						'on this case'
					).format(sc_name, self.base_sc_folder)
				)
				self.df_sc[name, constants.Results.skipped] = True
				continue
			else:
				# Get first reference
				pf_sc = pf_sc[0]

			# Find handle in powerfactory for operating scenario
			pf_os = self.base_os_folder.GetContents(os_name)
			if len(pf_os) == 0:
				# Study case doesn't exist so alert user and skip to next
				self.logger.error(
					(
						'Operating Scenario {} cannot be found in PowerFactory folder {}, no studies will be carried out '
						'on this case'
					).format(os_name, self.base_os_folder)
				)
				self.df_sc[name, constants.Results.skipped] = True
				continue
			else:
				# Get first reference
				pf_os = pf_os[0]

			# Create a copy of these study_cases and scenarios
			new_sc, new_os = self.copy_study_case(name=name, sc=pf_sc, os=pf_os)

			# Create a PFStudyCase instance
			study_case_class = PFStudyCase(
				name=name, sc=new_sc, op=new_os, prj=self.prj,
				base_case=True,
				res_pth=self.res_pth
			)

			study_case_class.create_studies(lf_settings=self.lf_settings, fs_settings=self.fs_settings)
			# TODO: replace with a create_studies command
			# Assign relevant load flow to study case
			# study_case_class.create_load_flow(lf_settings=self.lf_settings)
			# # Assign relevant frequency scan to study case
			# study_case_class.create_freq_sweep(fs_settings=self.fs_settings)

			base_study_cases[name] = study_case_class

		return base_study_cases

	def copy_study_case(self, name, sc, os):
		"""
			Copy the study case and operating scenario to the temporary folders
		:param str name:  Name to use for new study case
		:param powerfactory.DataObject sc: Study case to be copied
		:param powerfactory.DataObject os: Scenario to be copied
		:return (powerfactory.DataObject, powerfactory.DataObject) (new_sc, new_op):  Handles to the newly created
																					study cases and operating scenarios
		"""
		# Ensure study case is deactivated before trying to copy
		self.deactivate_study_case()
		new_sc = self.sc_folder.AddCopy(sc, name)
		new_os = self.op_folder.AddCopy(os, name)

		if new_sc is None or new_os is None:
			self.logger.error(
				(
					'Unable to copy one of the following:\n\t'
					' - Study case {} to folder: {}\n\t'
					' - Operating Scenario {} to folder: {}\n'
					'Therefore no studies will be carried out on this case'
				).format(sc, self.sc_folder, os, self.op_folder)
			)
			self.df_sc.loc[name, constants.Results.skipped] = True

		return new_sc, new_os

	def deactivate_study_case(self):
		"""
			Deactivate the active study case
		:return None:
		"""
		# Get handle for active study case from PowerFactory
		study = app.GetActiveStudyCase()

		# If already deactivated then do nothing otherwise deactivate
		if study is not None:
			sce = study.Deactivate()
			if sce == 0:
				self.logger.debug('Deactivated active study case <{}> successfully'.format(study))
			elif sce > 0:
				self.logger.warning('Unable to deactivate study case <{}>, powerfactory return error code: {}'.format(
					study, sce
				)
				)
		return None

	# def process_fs_results(self, logger=None):
	# 	""" Loop through each study case cls and process results files
	# 	:return list fs_res
	# 	"""
	# 	fs_res = []
	# 	for sc_cls in self.sc_cases:
	# 		fs_res.extend(sc_cls.process_fs_results(logger=logger))
	# 	return fs_res

	# def process_hrlf_results(self, logger):
	# 	""" Loop through each study case cls and process results files
	# 	:return list hrlf_res:
	# 	"""
	# 	hrlf_res = []
	# 	for sc_cls in self.sc_cases:
	# 		hrlf_res.extend(sc_cls.process_hrlf_results(logger))
	# 	return hrlf_res

	def update_auto_exec(self):
		"""
			For the newly added study cases, updates the frequency sweep and adds to the auto_exec command
		:return None:
		"""
		for case in self.cases_to_run:
			self.task_auto.AppendStudyCase(case.sc)

			# Add frequency scan commands and results export
			self.task_auto.AppendCommand(case.fs, 0)
			self.task_auto.AppendCommand(case.fs_export_cmd, 0)

		return None

	def project_state(self, deactivate=False):
		"""
			Function to toggle the status of this project
		:return None:
		"""
		if deactivate:
			self.pf.deactivate_project()
			self.prj_active = False
		else:
			self.pf.activate_project(project_name=self.name)
			self.prj_active = True

		return None

	def create_folder(self, name, location, temp=False):
		"""
			Create temporary folders within the project to store newly created study cases, etc.
		:param str name:  Name to give to folder
		:param powerfactory.DataObject location:  Location new folder is to be created in
		:param bool temp: (optional) If True then marked as a temporary folder and added to the list of folders
		:return powerfactory.DataObject new_folder:  Handle to the newly created folder
		"""
		# Default location is in the project (assuming project has been activated successfully)
		if not self.prj_active:
			self.logger.critical(
				'Attempting to create folder {} in non-active project {}'.format(name, self.name)
			)
			raise SyntaxError('Creating folder in a non-active project')

		self.logger.debug('Creating new folder {} in location {}'.format(name, location))

		# In case name already has IntProject, remove from the name
		name = name.replace('.{}'.format(constants.PowerFactory.pf_folder_type),'')

		# Check if folder already exists and if so append (1) to name
		i = 0
		folder_exists = True
		new_name = name
		while folder_exists:
			# If on 2nd loop then append (i) to end of original name
			if i > 0:
				new_name = '{}({})'.format(name, i)
				self.logger.debug(
					'Original folder {} already exists in location {}, testing folder name {}'.format(
						name, location, new_name)
				)

			existing_folder = location.GetContents('{}.{}'.format(new_name, constants.PowerFactory.pf_folder_type))
			if len(existing_folder) > 0:
				folder_exists = True
			else:
				folder_exists = False
			i+=1

		# Alert user to change
		if new_name != name:
			self.logger.warning(
				(
					'Folder name {} already existed in PowerFactory project <{}>, new folder {} created instead'
				).format(name, location, new_name)
			)

		# Create folder
		new_folder = location.CreateObject(constants.PowerFactory.pf_folder_type, new_name)
		if new_folder is None:
			self.logger.error(
				'Unable to create folder {} in location {}, the script is likely to now fail'.format(name, location)
			)
		else:
			self.logger.debug('New folder {} created in location {}'.format(new_name, location))

		# If a temporary folder then add to list of temporary folders
		if temp:
			self.temp_folders.append(new_folder)
			self.logger.debug('Folder: {} added to list of temporary folders for deletion at the end'.format(new_folder))

		return new_folder

	def delete_temp_folders(self):
		"""
			Routine to delete all of the temporary folders created initially
		:return None:
		"""

		for folder in self.temp_folders:
			if folder is not None:
				self.pf.delete_object(pf_obj=folder)
				self.logger.debug('Temporary folder {} deleted'.format(folder))
			# folder = None

		self.logger.debug('Temporary folders created in project {} have all been deleted'.format(self.prj))

		return None

	def find_substation(self, sub_name):
		"""
			Function searches relevant possible locations that a substation could be located and returns
			the substation or an error message when multiple found
		:param str sub_name:  Name of substation to be found
		:return powerfactory.DataObject substation: Reference to the powerfactory substation element
		"""
		# Check ends with the substation element ending
		if not sub_name.endswith(constants.PowerFactory.pf_substation):
			sub_name = '{}.{}'.format(sub_name, constants.PowerFactory.pf_substation)

		# Find substation using a recursive search of the network elements folders
		substation = list()
		for net_item in self.net_data_items:
			# Loop through each net_item folder and extend substation
			substation.extend(net_item.GetContents(sub_name))

		# Check that only a single substation is found
		if len(substation) == 0:
			substation = None
		elif len(substation) > 1:
			self.logger.error(
				(
					'Multiple substations with the name {} have been found across multiple network data folders.'
					'The following substations where found: \n\t'
					'{}\n'
				).format(sub_name, '\n\t'.join([str(x) for x in substation]))
			)
			substation = None
		else:
			substation = substation[0]

		return substation

	def create_fault_cases(self, contingencies):
		"""
			Function will loop through all of the contingencies and create a fault case for each which are
			all added to a temporary folder.
			This list of fault cases can then be added to a contingency case and each study case / operating scenario
			associated with a project tested for convergence.
		:param dict contingencies:  Reference to the contingencies returned in a dictionary as part of the inputs
									processing
		:return dict fault_cases:  Returns a dictionary which contains a reference to all of the fault cases created
		"""
		# Fault cases list initialised to be empty
		fault_cases = dict()

		# Find base folder for all fault cases to be stored in
		faults_folder = app.GetProjectFolder(constants.PowerFactory.pf_faults_folder_type)


		# Create temporary folder to store all of the fault cases within and add to list of folders to be deleted
		# self.fault_case_folder = self.create_folder(
		# 	name='{}_{}'.format(constants.PowerFactory.temp_faults_folder, constants.uid),
		# 	location=faults_folder
		# )
		self.fault_case_folder, _ = create_object(
			location=faults_folder,
			pfclass=constants.PowerFactory.pf_fault_cases_folder,
			name='{}_{}'.format(constants.PowerFactory.temp_faults_folder, constants.uid)
		)
		self.temp_folders.append(self.fault_case_folder)

		# Loop through each contingency and look for relevant elements
		for name, cont in contingencies.items():
			# Check if status of contingency is set to skip and if so continue
			if cont.skip:
				self.logger.debug(
					'Contingency {} is not considered for analysis and is therefore skipped'.format(cont.name)
				)
				continue

			# Create new switch event within the network folder
			fault_event, _ = create_object(
				location=self.fault_case_folder,
				pfclass=constants.PowerFactory.pf_fault_event,
				name=cont.name
			)

			# Assign as a contingency case
			fault_event.mod_cnt = 1

			# Loop through each coupler and add switch event to fault case
			for coupler in cont.couplers:
				# Find substation using a recursive search of the network elements folders
				substation = self.find_substation(sub_name=coupler.substation)

				if substation is None:
					# Not able to find substation and therefore contingency cannot be found
					self.logger.error(
						(
							'For contingency {} the substation named {} cannot be found within the project '
							'{} and therefore the contingency will not be studied.'
						).format(cont.name, coupler.substation, self.prj)
					)
					break

				# Find the switch within this substation
				breaker = substation.GetContents(coupler.breaker)

				# Check that only a single substation is found
				if len(breaker) == 0:
					self.logger.error(
						(
							'For contingency {} the circuit breaker named {} cannot be found within the substation'
							'<{}> and therefore this contingency will not be studied'
						).format(cont.name, coupler.breaker, substation)
					)
					cont.not_included = True
					break
				else:
					breaker = breaker[0]

				switch_event, _ = create_object(
					location=fault_event,
					pfclass=constants.PowerFactory.pf_switch_event,
					name=breaker.loc_name
				)
				# Set target element
				switch_event.p_target = breaker
				# Set status and ensure takes place on all phases
				switch_event.i_switch = coupler.status
				switch_event.i_allph = 1

			# Check if all events added successfully otherwise delete fault case
			if cont.not_included:
				fault_event.Delete()
			else:
				self.logger.debug('Fault case {} successfully created for contingency {}'.format(fault_event, cont.name))
				fault_cases[cont.name] = fault_event
				# Reference to the created fault event added to the contingency record
				cont.fault_event = fault_event

		# Populate dictionary with Fault Cases
		self.fault_cases = fault_cases
		return fault_cases

	def pre_case_check(self, contingencies=None, contingencies_cmd=str()):
		"""
			Function runs through all of the base study cases and checks which contingencies pass
			the user is then provided with a dataframe summarising for this project all of the study case
			and operating scenario combinations that pass

		:param dict contingencies:  (optional) Dictionary of the outages to be considered which will need to be
									created into fault cases
		:param str contingencies_cmd: (optional) String of the command to be used for contingency analysis
		:return pd.DataFrame df_status:  Combined DataFrame showing those which are convergent
		"""

		# Create fault cases if no command provided
		if not contingencies_cmd:
			if contingencies:
				self.create_fault_cases(contingencies=contingencies)
			else:
				self.logger.critical('No contingency command or dictionary of contingencies provided, '
									 'not possible to run analysis')
				raise SyntaxError('Incomplete inputs, missing contingencies and contingencies_cmd')
		else:
			if contingencies:
				self.logger.warning('Input provided for both contingencies and contingencies_cmd, '
									'contingencies will be used as preference')
				self.fault_cases = dict()

		# Loop through each of the base study cases, run the contingency analysis, process the results
		# and then combine the results into a single dataframe
		for sc_name, sc in self.base_sc.items():
			# Ensure study case is active in PowerFactory
			sc.toggle_state()

			# Check if sc has a contingency analysis defined
			if not sc.cont_analysis:
				# Contingency analysis function has not been defined so need to create but only possible if
				# fault cases have been or contingency command provided as an input.
				sc.create_cont_analysis(fault_cases=self.fault_cases, cmd=contingencies_cmd)

			# Run the contingency analysis for this study case
			sc.cont_analysis.Execute()

			# Process the results so that the DataFrame is up to date
			sc.process_cont_results()

		# Get a dictionary of all of the study_case DataFrames so they can be combined to the project
		# level
		dfs = {sc_name: sc.df for sc_name, sc in self.base_sc.items()}
		df = pd.concat(
			dfs.values(), axis=0, keys=dfs.keys(),
			names=(constants.Contingencies.sc, constants.Contingencies.cont))

		self.logger.debug('Pre case check results combined for project {}'.format(self.prj))

		# Assign the pre_case_check dataframe for this project
		# TODO: May be better to put this in the same as the df_sc DataFrame
		self.df_pre_case = df

		return df

	def create_cases(self, study_settings, terminals=None, contingencies=None, contingencies_cmd=str()):
		"""
			Function runs the pre_case_check for all of the base study_cases and then creates the study cases for each
			contingency.
		:param file_io.StudySettings study_settings: User selected settings for this input
		:param dict terminals: Dictionary of the terminals which need to be calculated
		:param dict contingencies:  (optional) Dictionary of the outages to be considered which will need to be
									created into fault cases
		:param str contingencies_cmd: (optional) String of the command to be used for contingency analysis
		:return None:
		"""
		# If pre_case_check has not yet been run then run now
		if self.df_pre_case.empty:
			self.logger.debug('Running pre-case check')
			_ = self.pre_case_check(contingencies=contingencies, contingencies_cmd=contingencies_cmd)

		# Check terminals have been defined otherwise do that now
		if not self.terminals:
			_ = self.find_terminals(terminals_to_include=terminals, include_mutual=study_settings.export_mutual)

		df_convergent = self.df_pre_case[self.df_pre_case[constants.Contingencies.status]==True]

		# Check if the intact case should be included and then if so add to cases
		self.cases_to_run = list()
		if study_settings.include_intact:
			# TODO: Produce test routine to confirm this works
			# Update export path and results files and then add study case to results
			for _, sc in self.base_sc.items():
				sc.res_pth = study_settings.export_folder
				sc.create_studies()
				self.cases_to_run.append(sc)
		else:
			self.cases_to_run = list()


		if df_convergent.empty:
			msg = 'No convergent contingencies found for cases in the project {}.\n'.format(self.prj)
			if study_settings.include_intact:
				self.logger.warning(
					'{} Results will be run for the following intact study cases only:\n\t{}'.format(
						msg, '\n\t'.join([sc.name for sc in self.cases_to_run])
					)
				)
			else:
				self.logger.warning(
					(
						'{} The user has decided not to include the intact system and therefore no results will be '
						'returned.'
					).format(msg)
				)

		else:
			# Loop through each study case to create new cases based on those and the relevant contingencies
			for sc_name, sc in self.base_sc.items():
				# Add the terminals to the results file for each of the base study cases before the new cases are
				# created which uses them as a starting point
				sc.add_variables(study_settings=study_settings, terminals=self.terminals, mutuals=self.mutuals)

				# Create cases for all the convergent contingencies associated with this study case and then returns
				# a list of references to the PFStudyCase class
				new_cases = sc.create_cases(
					sc_folder=self.sc_folder, op_folder=self.op_folder, res_pth=study_settings.export_folder
				)

				# Add details of newly created cases to the overall list
				self.cases_to_run.extend(new_cases)

		return None

	def create_task_auto(self):
		"""
			Function creates the command for automation of the study results and is saved in the temporary
			study case folder
		:return None:
		"""
		# Check if study case folder has been created and if not then create
		if not self.sc_folder:
			self.logger.warning('Temporary study case folder has not been created and so will be created now')
			self.sc_folder = self.create_folder(
				name='{}_{}'.format(constants.PowerFactory.temp_sc_folder, self.uid),
				location=self.base_sc_folder,
				temp=True
			)

		# Create the auto command
		task_auto, _ = create_object(
			location=self.sc_folder,
			pfclass=constants.PowerFactory.autotasks_command,
			name='{}_{}'.format(constants.General.cmd_autotasks_leader, self.uid)
		)
		return task_auto

	def find_terminals(self, terminals_to_include, include_mutual=False):
		"""
			Function finds all the terminals in the active project and returns details of those
			which cannot be found
		:param dict terminals_to_include:  List of terminals as defined in file_io.TerminalDetails
		:param bool include_mutual:  Set to True when mutual impedance values are supposed to be exported
		:return pd.DataFrame df_missing_terminal:  Returns details of all the terminals found in project
		"""
		self.logger.debug('Checking for relevant terminals in project:  {}'.format(self.prj))

		# Empty DataFrame which will be populated with the status of this terminal for this project
		c = constants.Terminals
		df = pd.DataFrame(columns=c.columns)

		# Confirm project is active
		# TODO: What happens if try to find a terminal that exists in project but not study case
		# TODO: What happens if no study case active
		self.project_state()

		# Input dictionary is duplicated since the pf_reference is project specific
		self.terminals = dict()
		# Loop through each terminal provided as an input and check if it can be found, if it can update the
		# terminal with the PowerFactory handle
		for term_name, terminal in terminals_to_include.items():
			self.logger.debug(
				(
					'Looking for terminal {}, associated with substation {} and busbar {} in project {}'
				).format(term_name, terminal.substation, terminal.terminal, self.prj)
			)
			# Populate DataFrame with details for this terminal
			df.loc[term_name, c.name] = terminal.name
			df.loc[term_name, c.sub1] = terminal.substation
			df.loc[term_name, c.bus1] = terminal.terminal
			df.loc[term_name, c.include_mutual] = terminal.include_mutual

			# Find substation which contains this terminal
			pf_sub = self.find_substation(sub_name=terminal.substation)

			if pf_sub is None:
				# Error message displayed at end for all terminals that cannot be found
				found = False

			else:
				# Check if terminal is contained within substation
				# Get list of all terminals that match this name
				terminals_in_substation = pf_sub.GetContents(terminal.terminal)

				# Confirm that at least 1 terminal with the required named exists in the substation

				if len(terminals_in_substation) == 0:
					# Error message displayed at end for all terminals that cannot be found
					found = False


				else:
					if len(terminals_in_substation) > 1:
						# If multiple terminals with the same name exist then alert User.  This should not be possible in the current
						# version of PowerFactory
						self.logger.warning(
							(
								'More than 1 terminal with the name {} found in substation {} for PowerFactory Project {} '
								'and this relates to Terminal Input {}.\n Results will only be returned for the 1st one of'
								'the following list of terminals found: \n\t {}'
							).format(terminal.terminal, terminal.substation, self.prj, terminal.name,
									 '\n\t'.join([str(x) for x in terminals_in_substation])
									 )
						)

					# Terminal found so create a reference t it
					found = True
					pf_terminal = terminals_in_substation[0]

					# Only those which exist are now available in this project
					new_term_object = file_io.TerminalDetails(
						name=term_name,
						substation=terminal.substation,
						terminal=terminal.terminal,
						include_mutual=terminal.include_mutual,
					) # type: file_io.TerminalDetails
					new_term_object.found = found
					new_term_object.pf_handle = pf_terminal
					# Create reference to terminal and then add to dictionary
					self.terminals[term_name] = new_term_object

			# Update terminals dictionary and DataFrame of status
			df.loc[term_name, c.status] = found

		# All terminals have been added so print list of terminals which couldn't be found as warning to user
		missing_terms = df[df[c.status]==True].index
		if len(missing_terms) != len(terminals_to_include):
			# Number of terminals expected does not match with number found
			self.logger.warning(
				(
					'The following terminals details in the inputs cannot be found in the project {}, no results'
					'will be returned for these terminals and so you may wish to check the inputs:\n\t{}'
				).format(self.prj, missing_terms)
			)
		else:
			# All terminals found
			self.logger.info('All input terminals found in project: {}'.format(self.prj))

		# Create mutual impedance elements and obtain updated DataFrame
		if include_mutual:
			df = self.create_mutual_impedance(df=df)
		else:
			self.logger.debug('No mutual impedance values requested for project {}'.format(self.prj))

		# Returns DataFrame with details of terminals that have been found and those which are missing
		return df

	def create_mutual_impedance(self, df):
		"""
			Based on the terminals that have been found within the project the mutual impedance elements are
			created and are located in the Network data folders.
			Mutual impedance elements have to be stored in the network data for the active project

		:param pd.DataFrame df:  DataFrame of terminals that have been found already, this is popualted further and
								returned
		:return pd.DataFrame, df:  Returns a DataFrame with the referencing for the mutual elements created
		"""

		self.logger.debug('Creating: Mutual Impedance Elements for project {}'.format(self.prj))
		c = constants.Terminals

		if not self.terminals:
			# If no terminals exist then no mutual impedance elements to create
			self.logger.warning(
				(
					'No terminals could be found for the project {} and therefore no mutual impedance elements '
					'could be created.'
				).format(self.prj)
			)
		else:

			# Create temporary folder to store the mutual impedance elements
			# Folder has to be in one of the network element folders for results to be calculated
			# TODO: Confirm if it matters which one
			mutual_folder = self.create_folder(
				name='{}_{}'.format(constants.PowerFactory.temp_mutual_folder, self.uid),
				location=self.net_data,
				temp=True
			)
			# For some reason cannot directly create in the required location so have to move after creation
			# object handle is updated automatically
			if mutual_folder is not None:
				self.net_data_items[0].Move(mutual_folder)
			else:
				self.logger.error(
					(
						'Unable to create a temporary folder for the mutual impedance elements in the location {} '
						'and therefore no mutual elements can be created'
					).format(self.net_data_items[0])
				)
				# Return eearly
				return df

			# Reset mutual elements dictionary which is populated for each mutual element created in the form
			# of having the name (term1_term2) and then the reference to the powerfactory DataObject that is created
			self.mutuals = dict()

			# Loop through all terminals that have already been found
			for name, term in self.terminals.items():
				if term.include_mutual:
					# Element is set to include mutual and therefore need to create a new mutual element
					# for every link from this terminal to another terminal
					for other_name, other_term in self.terminals.items():
						# Don't create mutual impedance to own terminal
						if other_name != name:
							planned_name, used_name = create_mutual_name(term1=name, term2=other_name)

							# Update dataframe
							df.loc[used_name, c.name] = used_name
							df.loc[used_name, c.sub1] = term.substation
							df.loc[used_name, c.bus1] = term.terminal
							df.loc[used_name, c.include_mutual] = term.include_mutual
							df.loc[used_name, c.planned_name] = planned_name
							df.loc[used_name, c.sub2] = other_term.substation
							df.loc[used_name, c.bus2] = other_term.terminal
							df.loc[used_name, c.status] = True

							# Create mutual element in the mutual folder
							elmmut = create_mutual_elm(
								location=mutual_folder,
								name=used_name,
								bus1=term.pf_handle,
								bus2=other_term.pf_handle
							)

							self.mutuals[used_name] = elmmut

							self.logger.debug(
								'Mutual impedance element {}, created between terminal {} and {}'.format(
									elmmut, term.pf_handle, other_term.pf_handle
								)
							)

		# Return updated DataFrame with mutual elements
		return df

	def run_parallel_tasks(self):
		"""
			Function to run parallel tasks and then detects if an error has occured.
			If an error occurs will run in non-parallel mode with a warning message to user
		:return None:
		"""
		self.logger.info(
			'Starting parallel running of studies for project {} using command {}'.format(
				self.prj, self.task_auto
			)
		)

		# Execute command
		ierr = self.task_auto.Execute()

		if ierr > 0:
			self.logger.warning(
				(
					'An error occurred trying to run the command {} on parallel processors, this could be'
					'either a licensing issue or a PowerFactory response delay.  The study will be attempted using'
					'non parallel processes'
				).format(self.task_auto)
			)
			# Change task_auto settings to disable use of parallel processing
			self.task_auto.iEnableParal = 0

			# Execute
			ierr = self.task_auto.Execute()

		if ierr > 0:
			self.logger.critical(
				(
					'Unable to run results for project {}, there may be a license issue that the user should look into. '
					'The script will now fail and all the following temporary folders will remain so that the user can '
					'investigate the issue more closely.\n\t{}'
				).format(self.prj, '\n\t'.join([str(x) for x in self.temp_folders]))
			)
			raise RuntimeError('Not able to run studies after multiple attempts')
		else:
			self.logger.info('Studies completed for project {}'.format(self.prj))


class PowerFactory:
	"""
		Class to deal with system level interfacing in PowerFactory
	"""
	# TODO: Check correct license exists

	def __init__(self):
		""" Gets the relevant powerfactory version and initialises """
		self.c = constants.PowerFactory
		self.logger = logging.getLogger(constants.logger_name)

		# Constants
		self.settings = None

	def add_python_paths(self):
		"""
			Function retrieves the relevant python paths, adds them and then imports the powerfactory module
			Importing of the powerfactory module has to happen here due to the location
		"""
		# Get the python paths if not already populated
		if not (self.c.dig_path and self.c.dig_python_path):
			# Initialise so that the paths are looked for
			self.c = self.c()

		# Add the paths to system and the environment and then try and import powerfactory
		sys.path.append(self.c.dig_path)
		sys.path.append(self.c.dig_python_path)
		os.environ['PATH'] = '{};{}'.format(os.environ['PATH'], self.c.dig_path)

		# Try and import the powerfactory module
		try:
			global powerfactory
			import powerfactory
		except ImportError:
			self.logger.critical(
				(
					'It has not been possible to import the powerfactory module and therefore the script cannot be run.\n'
					'This is most likely due to there not being a powerfactory.pyc file located in the following path:\n\t'
					'<{}>\n'
					'Please check this exists and the error messages above.'
				).format(self.c.dig_python_path)
			)
			raise ImportError('PowerFactory module not found')

		return None

	def initialise_power_factory(self):
		"""
			Function initialises powerfactory and provides an object reference to it
		:return None:
		"""
		# Check the paths have already been found and if not call the relevant function
		if not (self.c.dig_path and self.c.dig_python_path):
			# Initialise so that the paths are looked for
			self.c = self.c()
			self.add_python_paths()

		# Different APIs exist for different PowerFactory versions, if an old version is run then different
		# initialisation route.  When initialising need to warn user that old version is being used
		global app
		# Only initialise PowerFactory if not already initialised
		if app is None:
			if distutils.version.StrictVersion(powerfactory.__version__) > distutils.version.StrictVersion('17.0.0'):
				# Error sometimes in getting access to a license which returns certain error codes and therefore
				# script will now make a few attempts for those cases
				app_init_count = 0
				while app is None:
					# powerfactory after 2017 has an error handler when trying to load
					app_init_count += 1
					self.logger.debug('Attempting license activation number: {}'.format(app_init_count))
					try:
						app = powerfactory.GetApplicationExt()  # Start PowerFactory  in engine mode
					except powerfactory.ExitError as error:
						# Will attempt to connect to license upto this many times
						if (
								app_init_count < constants.PowerFactory.license_activation_attempts and
								error.code in constants.PowerFactory.license_activation_error_codes
						):
							self.logger.warning(
								(
									'Unable to connect to license due to a license error which returned error message'
									'\n\t{}\n and associated error code: {}.\n'
									'This may be an intermittent error and the script will try again in {:.0f} seconds.\n'
									'This will be attempt {} of {}'
								).format(
									error,
									error.code,
									constants.PowerFactory.license_activation_delay,
									app_init_count,
									constants.PowerFactory.license_activation_attempts)
							)
							time.sleep(constants.PowerFactory.license_activation_delay)

						elif app_init_count >= constants.PowerFactory.license_activation_attempts:
							# A certain number of attempts have been made and now is the time to fail
							self.logger.critical(
								(
									'Have attempted to connect to PowerFactory {} times and still receiving an error \n\t{}\n '
									'with associated error code {}.  There is likely to be some permanent connecting to '
									'the license that the user will need to look into!'
								).format(app_init_count, error, error.code)
							)
							raise ImportError('Unable to load PowerFactory due to a license issue - Unable to run')

						else:
							# A different error code has been returned and so fail the script
							self.logger.critical(
								(
									'An error occurred trying to start PowerFactory.\n'
									'The following error message was returned by PowerFactory \n\t{}\n'
									'and associated error code: {}'
								).format(error, error.code)
							)
							raise ImportError('Power Factory Load Error - Unable to run')

			else:
				# In case of an older version of PowerFactory being run
				app = powerfactory.GetApplication()
				if app is None:
					self.logger.critical(
						'Unable to load PowerFactory and this version does not return any error codes, you will need '
						'to user a newer version of PowerFactory or investigate the error messages detailed above.'
					)
					raise ImportError('Power Factory Load Error - Unable to run')

				# Clear the powerfactory output window
				app.ClearOutputWindow()  # Clear Output Window

		# Call function to confirm that the PQ license is available otherwise fail script
		self.check_pq_license_exists()

		return None

	def activate_project(self, project_name):
		"""
			Activate a project for which a name is provided and return False if project cannot be found
		:param str project_name:  Name of project to be activated
		:return powerfactory.DataObject pf_prj: Either returns a handle to the project or False if fails
		"""
		# Before trying to activate a project confirm that PowerFactory has been initialised
		if not app:
			self.initialise_power_factory()

		# Check project name already has correct ending and if not add
		if not project_name.endswith(self.c.pf_project):
			project_name = '{}.{}'.format(project_name, self.c.pf_project)

		success = app.ActivateProject(project_name)

		# If successfully activated then get a handle for this project
		if not success:
			pf_prj = app.GetActiveProject()
		else:
			pf_prj = None

		return pf_prj

	def get_active_project(self):
		"""
			Function just returns a handle to the currently active project
		:return powerfactory.DataObject pf_prj:
		"""
		# Get reference to the currently activated project
		pf_prj = app.GetActiveProject()

		return pf_prj

	def import_project(self, project_pth):
		"""
			This function will import a project into PowerFactory and then activates
			Project is imported to the current user location
			Further details here:
				https://www.digsilent.de/en/faq-reader-powerfactory/how-to-import-export-pfd-files-by-using-script.html
		:param str project_pth:  Path to the project file to be imported
		:return None: Will throw errors if not possible to import
		"""

		# Check file exists before continuing
		if not os.path.exists(project_pth):
			self.logger.critical(
				(
					'The following file containing the PowerFactory project to be imported does not exist:\n\t{}'
				).format(project_pth)
			)
			raise ValueError('No file exists to import')

		# Determine the location the imported project should be saved which is the current user
		location = app.GetCurrentUser()

		# Create an object for the import command
		import_object = location.CreateObject('CompfdImport', 'Import')

		# Set the file attribute to be imported and the target location (current user)
		import_object.SetAttribute('e:g_file', project_pth)
		import_object.g_target = location

		# Execute command (returns 0 for success) and alert user if error importing
		error = import_object.Execute()

		if error:
			self.logger.critical(
				(
					'Critical error occurred when trying to import the project from location:\n\t{}\n'
					'PowerFactory returned the following error code from the function\n'
					'\tCommand: {}\n'
					'\tTarget User Location: {}\n'
					'\tError Code: {}'
				).format(project_pth, import_object, location, error)
			)
			raise ImportError('Unable to import the PowerFactory project')

		# Delete newly created script
		self.delete_object(import_object)

		return None

	def deactivate_project(self):
		""" Function to deactivate the current project """
		pf_prj = self.get_active_project()

		# Deactivate project if project active
		if pf_prj:
			error = pf_prj.Deactivate()
		else:
			error = 0

		if error:
			self.logger.error(
				(
					'Unable to deactivate the currently active project {}'
				).format(pf_prj.GetFullName(type=0))
			)

		return None

	def delete_object(self, pf_obj):
		"""
			Function will delete a PowerFactory object from the DataManager
		:param powerfactory.DataObject pf_obj:  Reference to the object to be deleted
		:return None:
		"""

		# Function deletes the object (it is only moved to the Recycle Bin)
		error = pf_obj.Delete()

		if error:
			self.logger.error(
				(
					'It has not been possible to delete the following object\n\t{}'
				).format(pf_obj.GetFullName(type=0))
			)

		return None

	def check_pq_license_exists(self):
		"""
			Check that the activated power factory installation has a license for the PowerQuality module otherwise
			there are no studies that can be run
		:return None:
		"""
		# Get reference to the current user
		pf_user = app.GetCurrentUser()

		# Get license status for power quality module and alert user if not suitable
		harm_license = pf_user.harm
		if harm_license == 0:
			self.logger.critical(
				(
					'The activated PowerFactory does not have access to the power quality module required to run '
					'frequency scans.  Therefore all this study will be able to do anything.'
					'\n\t - PowerFactory Power Quality and Harmonic Module = {}\n'
					'This script is running PowerFactory {} and opening that version will allow you to enable the '
					'license in the UserManager'
				).format(harm_license, self.c.pf_year)
			)
			raise EnvironmentError('PowerFactory Power Quality License Not Available')

		return None

	# def check_parallel_processing(self):
	# 	"""
	# 		Function determines the number of processes that powerfactory is set to run
	# 	"""
	# 	# TODO: Requires reference to ParallelMan Settings to work, needs further development
	# 	# NOT CURRENTLY WORKING
	#
	# 	# Get number of cpus available
	# 	number_of_cpu = multiprocessing.cpu_count()
	#
	# 	# Check number of processors set to be run
	# 	current_processors = app.GetNumSlave()
	#
	# 	# Display warning of a small value
	# 	if current_processors == 1 or current_processors < (number_of_cpu-1):
	# 		self.logger.warning(
	# 			(
	# 				'Your PowerFactory settings are set to only allow running on {} parallel processors, this does not'
	# 				'take full advantage of the machines capability which has {} processors and therefore may take '
	# 				'longer to run.'
	# 			).format(current_processors, number_of_cpu)
	# 		)
	#
	# 	return None

	def change_parallel_settings(self, delay=constants.PowerFactory.parallel_time_out, reduce=False):
		"""
			Function will change some of the parallel processing settings to increase the time allowed
			before a response is necessary
		:param int delay: Maximum delay when waiting for parallel processor response
		:param bool reduce: (optional) If set to True then will reduce as well as increase
		:return int existing_delay: Returns the original delay value incase needs restoring
		"""
		# Before trying to activate a project confirm that PowerFactory has been initialised
		if not app:
			self.initialise_power_factory()

		# Get reference to current user
		current_user = app.GetCurrentUser()

		# Get the default settings folder
		settings = current_user.GetContents(constants.PowerFactory.user_default_settings)

		if len(settings) == 0:
			self.logger.critical(
				(
					'Not able to find the default settings named {} in the current user {} and therefore not able to '
					'change the user settings'
				).format(constants.PowerFactory.user_default_settings, current_user)
			)
			raise EnvironmentError('Not able to find user default settings for which change is requested')
		else:
			# Get first element
			self.settings = settings[0]

		existing_delay = self.settings.procTimeOut

		if existing_delay < delay:
			self.logger.warning(
				(
					'The existing delay while waiting for parallel processes to return is {:.0f} seconds when the '
					'recommended value is at least {:.0f} seconds.  The settings value stored in {} has therefore '
					'been increased.'
				).format(existing_delay, delay, self.settings)
			)

			# Change delay value
			self.settings.procTimeOut = delay
		else:
			if reduce:
				self.logger.info(
					(
						'The existing delay to wait for parallel processes to return is {:.0f} seconds and the '
						'desired delay value is {:.0f} seconds.  Therefore the delay has been changed in the '
						'settings file {}'
					).format(existing_delay, delay, self.settings)
				)

				# Change delay value
				self.settings.procTimeOut = delay
			else:
				self.logger.debug(
					'The existing parallel processing delay is {:.0f} seconds and no changes have been made'.format(
						existing_delay
					)
				)



		return existing_delay


def create_pf_project_instances(df_study_cases, uid=constants.uid, lf_settings=None, fs_settings=None, export_pth=str()):
	"""
		Loops through each of the projects in the DataFrame of study cases and activates them to check they work
	:param pd.DataFrame df_study_cases:
	:return dict pf_projects:  Returns a dictionary of PF project instances
	"""
	# Loop through each project and create a PFProject instance, then check can activate
	pf_projects = dict()
	for project, df in df_study_cases.groupby(by=constants.StudySettings.project, axis=0):

		pf_project = PFProject(
			name=project, df_studycases=df, uid=uid, lf_settings=lf_settings, fs_settings=fs_settings,
			res_pth=export_pth
		)
		pf_projects[pf_project.name] = pf_project

	return pf_projects

def run_pre_case_checks(
		pf_projects, terminals, include_mutual=False, export_pth=str(), contingencies=None, contingencies_cmd=str(),

):
	"""
		Loop through each project so that it returns a DataFrame of all the study case results.

		If an export_pth is provided then these are also written to the target excel file
	:param dict pf_projects:  Dictionary of references to all projects being studied as returned by
							create_pf_project_instances
	:param dict terminals:  Dictionary of the terminals for which results need to be run
	:param bool include_mutual:  (optional) Set to True if mutual impedance data is to be exported
	:param str export_pth:  (optional) Export path to write results to if provided
	:param dict contingencies:  (optional) Dictionary of the outages to be considered which will need to be
									created into fault cases
	:param str contingencies_cmd: (optional) String of the command to be used for contingency analysis
	:return pd.DataFrame df_case_check: DataFrame showing contingencies which are convergent
	"""
	logger = logging.getLogger(constants.logger_name)
	c = constants.Contingencies

	# Loops through all projects and obtains DataFrame, these are then combined into a single DataFrame
	# ready to be written to excel
	dfs_cont = list()
	dfs_term = dict()
	for project_name, prj in pf_projects.items():
		# Activate project
		prj.project_state()

		# Obtain contingency analysis results for all relevant cases in this project
		df_cont = prj.pre_case_check(contingencies=contingencies, contingencies_cmd=contingencies_cmd)
		dfs_cont.append(df_cont)

		# Look for terminals in project and get DataFrame of those which cannot be found
		df_term = prj.find_terminals(terminals_to_include=terminals, include_mutual=include_mutual)
		dfs_term[project_name] = df_term

	# Combine returned DataFrames into a single DataFrame
	df_case_check_cont = pd.concat(dfs_cont)
	df_case_check_term = pd.concat(dfs_term.values(), keys=dfs_term.keys())

	# Loop through and detail those cases that are non-convergent
	df_non_conv = df_case_check_cont[df_case_check_cont[c.status]==False]
	if df_non_conv.empty:
		logger.info('All cases and contingencies convergent')
	else:
		msgs = list()
		for name, values in df_non_conv.iterrows():
			# Create a unique message for each combination
			# Name is a tuple of the study_case name and the contingency
			msgs.append(
				'{}: \t Project: {}, Study Case: {}, Operating Scenarios: {}, Contingency: {}'.format(
					'-'.join(name), values[c.prj], values[c.sc], values[c.op], values[c.cont]
				)
			)
		logger.warning(
			(
				'The following project, study case, operating scenario, contingency combinations were '
				'non-convergent during the pre-case check and therefore it will not be possible to run '
				'frequency scans for these cases.\n\t{}'
			).format('\n\t'.join(msgs))
		)

	# If a path has been provided then write it to excel
	if export_pth:
		with pd.ExcelWriter(export_pth) as xl:
			df_case_check_cont.to_excel(xl, sheet_name=constants.Contingencies.export_sheet_name)
			df_case_check_term.to_excel(xl, sheet_name=constants.Terminals.export_sheet_name)
		logger.info(
			'A summary status for all of the pre-case check results has been saved to the file: {}'.format(
				export_pth
			)
		)

	# Return the summary DataFrame
	return df_case_check_cont, df_case_check_term

def run_studies(pf_projects, inputs):
	"""
		Function runs the studies to create the cases and run all studies based on
		the provided dictionary of projects and input settings
	:param dict pf_projects:  Dictionary of projects for which all studies will be run
	:param file_io.StudyInputsDev inputs:  Input settings
	:return None
	"""
	t0 = time.time()
	logger = logging.getLogger(constants.logger_name)
	# Iterate through each project and create the various cases, the includes running a pre-case check but no
	# output is saved at this point
	for project_name, project in pf_projects.items():
		logger.debug('Studies being run for project {}:\t{}'.format(project_name, project.prj))
		project.create_cases(
			study_settings=inputs.settings,
			terminals=inputs.terminals,
			contingencies=inputs.contingencies,
			contingencies_cmd=inputs.contingency_cmd
		)

		logger.debug('Cases created for project: {}:\t{}'.format(project_name, project.prj))

		# Update the auto executable for this project
		project.update_auto_exec()

		# Batch run the results
		logger.info('Running of studies associated with project {} started'.format(project_name))
		project.run_parallel_tasks()
		t1 = time.time()
		logger.info('Running of studies associated with project {} completed in {:.0f} seconds'.format(
			project_name, t1-t0)
		)

		# Delete temporary folders created for this project
		if inputs.settings.delete_created_folders:
			project.delete_temp_folders()
		else:
			logger.info(
				(
					'As per user inputs, temporary folders associated with project {} have not been deleted and so will '
					'need tidying within PowerFactory directly'
				).format(project_name)
			)

	return None




