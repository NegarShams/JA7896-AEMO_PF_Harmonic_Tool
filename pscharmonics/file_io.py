"""
#######################################################################################################################
###											Excel Writing															###
###		Script deals with writing of data to excel and ensuring that a new instance of excel is used for processing	###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###		project JI6973 for EirGrid project PSPF010 - Specialise Support in Power Quality Analysis during 2018		###
###																													###
#######################################################################################################################
"""

import unittest
import os
import time
import numpy as np
import itertools
import pscharmonics.constants as constants
import logging
import pandas as pd


# Meta Data
__author__ = 'David Mills'
__version__ = '2.1.6'
__email__ = 'david.mills@pscconsulting.com'
__phone__ = '+44 7899 984158'
__status__ = 'In Development - Beta'

""" Following functions are used purely for processing the inputs of the HAST worksheet """
def add_contingency(row_data):
	"""
		Function to read in the contingency data and save to list
	:param list row_data:
	:return list combined_entry:
	"""
	if len(row_data) > 2:
		aa = row_data[1:]
		station_name = aa[0::3]
		breaker_name = aa[1::3]
		breaker_status = aa[2::3]
		breaker_name1 = ['{}.{}'.format(nam, constants.PowerFactory.pf_coupler) for nam in breaker_name]
		combined_entry = list(zip(station_name, breaker_name1, breaker_status))
		combined_entry.insert(0, row_data[0])
	else:
		combined_entry = [row_data[0], [0]]
	return combined_entry

def add_scenarios(row_data):
	"""
		Function to read in the scenario data and save to list
	:param list row_data:
	:return list combined_entry:
	"""
	combined_entry = [
		row_data[0],
		row_data[1],
		'{}.{}'.format(row_data[2], constants.PowerFactory.pf_case),
		'{}.{}'.format(row_data[3], constants.PowerFactory.pf_scenario)]
	return combined_entry

def add_terminals(row_data):
	"""
		Function to read in the terminals data and save to list
	:param tuple row_data: Single row of data from excel workbook to be imported
	:return list combined_entry: List of data as a combined entry
	"""
	logger = logging.getLogger(constants.logger_name)
	if len(row_data) < 4:
		# If row_data is less than 4 then it means an old HAST inputs sheet has probable been used and so a default
		# value will be assumed instead
		logger.warning(('No status given for whether mutual impedance should be included for terminal {} and '
						'so default value of {} assumed.  If this has happened for every node then it may be because '
						'an old HAST Input format has been used.')
					   .format(row_data[0], constants.HASTInputs.default_include_mutual))
		row_data = list(row_data) + [constants.HASTInputs.default_include_mutual]

	combined_entry = [
		row_data[0],
		'{}.{}'.format(row_data[1], constants.PowerFactory.pf_substation),
		'{}.{}'.format(row_data[2], constants.PowerFactory.pf_terminal),
		# Third column now contains TRUE or FALSE.  If True then data will be included including
		# transfer impedance from other nodes to this node.  If False then no data will be included.
		row_data[3]]

	return combined_entry

def add_lf_settings(row_data):
	"""
		Function to read in the load flow settings and save to list
	:param list row_data:
	:return list combined_entry:
	"""
	z = row_data

	# Input data adjusted for backwards compatibility prior to inclusion of automatic tap adjustment of phase shifters
	# in PowerFactory 2018
	if len(row_data) == 56:
		row_data.insert(4, constants.HASTInputs.def_automatic_pst_tap)

	# Convert a nan value to an empty string for backwards compatibility if no Load Flow Command has been provided
	# as an input
	if z[0] is np.nan:
		z[0] = str()

	combined_entry = [
		z[0],
		int(z[1]), int(z[2]), int(z[3]), int(z[4]), int(z[5]), int(z[6]), int(z[7]), int(z[8]), int(z[9]), int(z[10]),
		float(z[11]), int(z[12]), int(z[13]), int(z[14]), z[15], z[16], int(z[17]), int(z[18]), int(z[19]),	int(z[20]),
		float(z[21]), int(z[22]), float(z[23]), int(z[24]), int(z[25]), int(z[26]), int(z[27]),	int(z[28]), int(z[29]),
		int(z[30]), z[31], z[32], int(z[33]), z[34], int(z[35]), int(z[36]),
		int(z[37]), int(z[38]), int(z[39]), z[40], z[41], z[42], z[43], int(z[44]), z[45], z[46], z[47],
		z[48], z[49], z[50], z[51], z[52], z[53], int(z[54]), int(z[55]), int(z[56])]

	return combined_entry

def add_freq_settings(row_data):
	"""
		Function to read in the frequency sweep settings and save to list
	:param list row_data:
	:return list combined_entry:
	"""
	z = row_data
	combined_entry = [z[0], z[1], int(z[2]), z[3], z[4], z[5], int(z[6]), z[7], z[8], z[9],
					  z[10], z[11], z[12], z[13], int(z[14]), int(z[15])]
	return combined_entry

def add_hlf_settings(row_data):
	"""
		Function to read in the harmonic load flow settings and save to list
	:param list row_data:
	:return list combined_entry:
	"""
	z = row_data
	combined_entry = [int(z[0]), int(z[1]), int(z[2]), int(z[3]), z[4], z[5], z[6], z[7],
					  z[8], int(z[9]), int(z[10]), int(z[11]), int(z[12]), int(z[13]), int(z[14])]
	return combined_entry

def add_study_settings(row_data):
	"""
		Function to read in the study settings and save to list
		TODO: Convert this to a class / process data formats whilst in DataFrame
	:param list row_data:
	:return list combined_entry:
	"""
	z = row_data
	# No longer backwards compatible since the input settings for selecting a defined load flow file have been added
	combined_entry = [
		z[0], z[1], z[2], z[3], z[4], z[5], z[6], z[7],
		bool(z[8]), bool(z[9]), bool(z[10]), bool(z[11]), bool(z[12]), bool(z[13]), bool(z[14]), bool(z[15]),
		bool(z[16]), bool(z[17]), bool(z[18]), bool(z[19])
	]

	return combined_entry


class Excel:
	""" Class to deal with the writing and reading from excel and therefore allowing unittesting of the
	functions

	Note:  Each call will create a new excel instance
	"""
	def __init__(self, print_info=print, print_error=print):
		"""
			Function initialises a new instance of an excel application
		TODO: To be replaced with logging handler
		:param builtin_function_or_method print_info:  Handle for print engine used for printing info messages
		:param builtin_function_or_method print_error:  Handle for print engine used for printing error messages
		"""
		# Constants
		# Sheets and starting rows for analysis
		self.analysis_sheets = constants.analysis_sheets
		# IEC limits
		# If on input spreadsheet then can be used to test against allocated limits
		self.iec_limits = constants.iec_limits
		self.limits = list(zip(*[self.iec_limits[1]]))

		# Updated with logging handlers once setup finished
		self.log_info = print_info
		self.log_error = print_error

		# Value set to true if importing workbook is a success
		self.import_success = False

	def import_excel_harmonic_inputs(self, pth_workbook):  # Import Excel Harmonic Input Settings
		"""
			Import Excel Harmonic Input Settings
		:param str pth_workbook: Name of workbook to be imported
		:return analysis_dict: Dictionary of the settings for the analysis work
		"""
		# Initialise empty dictionary
		analysis_dict = dict()

		# #wb = self.xl.Workbooks.Open(pth_workbook)  # Open workbook
		# #c = self.excel_constants
		# print(c.xlDown)
		# #self.xl.Visible = False  # Make excel Visible
		# #self.xl.DisplayAlerts = False  # Don't Display Alerts

		# Loop through each worksheet defined in <analysis_sheets>
		for x in self.analysis_sheets:
			# Import worksheet using pandas
			# Import data for this specific worksheet
			# TODO: Need to add something to capture when sheet name is missing
			current_worksheet = x[0]
			df = pd.read_excel(
				pth_workbook, sheet_name=x[0], header=x[1], usecols=x[2], skiprows=x[3]
			)

			row_input = []
			# #current_worksheet = x[0]

			# Code only to be executed for these sheets
			if current_worksheet in constants.PowerFactory.HAST_Input_Scenario_Sheets:
				# Loop through each row of DataFrame and convert series to list
				for index, data_series in df.iterrows():
					# Convert the Series to a list and return the required data
					# drop nan values from series
					# TODO: Should include data checking on import to improve processing speed
					data_to_process = data_series.dropna()
					row_value = data_to_process.tolist()
					if current_worksheet == constants.PowerFactory.sht_Contingencies:
						row_value = add_contingency(row_data=row_value)

					# Routine for Base_Scenarios worksheet
					elif current_worksheet == constants.PowerFactory.sht_Scenarios:
						row_value = add_scenarios(row_data=row_value)

					# Routine for Terminals worksheet
					elif current_worksheet == constants.PowerFactory.sht_Terminals:
						row_value = add_terminals(row_data=row_value)

					# Routine for Filters worksheet
					elif current_worksheet == constants.PowerFactory.sht_Filters:
						row_value = FilterDetails(row_data=row_value)

					row_input.append(row_value)

			# More efficiently checking which worksheet looking at
			elif current_worksheet in constants.PowerFactory.HAST_Input_Settings_Sheets:
				# For these worksheets input settings are in a series and can be converted to a list directly
				row_input = df.iloc[:,0].tolist()

				if current_worksheet == constants.PowerFactory.sht_LF:
					# Process inputs for Loadflow_Settings
					row_input = add_lf_settings(row_data=row_input)

				elif current_worksheet == constants.PowerFactory.sht_Freq:
					# Process inputs for Frequency_Sweep settings
					row_input = add_freq_settings(row_data=row_input)

				elif current_worksheet == constants.PowerFactory.sht_HLF:
					# Process inputs for Harmonic_Loadflow
					row_input = add_hlf_settings(row_data=row_input)

				elif current_worksheet ==  constants.PowerFactory.sht_Study:
					# Process inputs for Study Settings
					row_input = add_study_settings(row_data=row_input)

			# Combine imported results in a dictionary relevant to the worksheet that has been imported
			analysis_dict[current_worksheet] = row_input  # Imports range of values into a list of lists

		# #wb.Close()  # Close Workbook
		return analysis_dict

class StudyInputs:
	"""
		Class used to import the Settings from the Input Spreadsheet and convert into a format usable elsewhere
	"""
	def __init__(self, hast_inputs=None, uid_time=time.strftime('%y_%m_%d_%H_%M_%S'), filename=''):
		"""
			Initialises the settings based on the HAST Study Settings spreadsheet
		:param dict hast_inputs:  Dictionary of input data returned from file_io.Excel.import_excel_harmonic_inputs
		:param str uid_time:  Time string to use as the uid for these files
		:param str filename:  Filename of the HAST Inputs file used from which this data is extracted
		"""
		c = constants.PowerFactory
		# General constants
		self.filename=filename

		self.uid = uid_time

		# Attribute definitions (study settings)
		self.pth_results_folder = str()
		self.results_name = str()
		self.progress_log_name = str()
		self.error_log_name = str()
		self.debug_log_name = str()
		self.pth_results_folder_temp = str()
		self.pf_netelm = str()
		self.pf_mutelm = str()
		self.pf_resfolder = str()
		self.pf_opscen_folder = str()
		self.pre_case_check = bool()
		self.fs_sim = bool()
		self.hrm_sim = bool()
		self.skip_failed_lf = bool()
		self.del_created_folders = bool()
		self.export_to_excel = bool()
		self.excel_visible = bool()
		self.include_rx = bool()
		self.include_convex_hull = bool()
		self.export_z = bool()
		self.export_z12 = bool()
		self.export_hrm = bool()

		# Attribute definitions (study_case_details)
		self.sc_details = dict()
		self.sc_names = list()

		# Attribute definitions (contingency_details)
		self.cont_details = dict()
		self.cont_names = list()

		# Attribute definitions (terminals)
		self.list_of_terms = list()
		self.dict_of_terms = dict()

		# Attribute definitions (filters)
		self.list_of_filters = list()

		# Load Flow Settings
		# Will contain full string to load flow command to be used
		self.pf_loadflow_command = str()
		# Will contain reference to LFSettings which contains all settings
		self.lf = LFSettings()

		# Process study settings
		self.study_settings(hast_inputs[c.sht_Study])

		# Process load flow settings
		self.load_flow_settings(hast_inputs[c.sht_LF])

		# Process List of Terminals
		self.process_terminals(hast_inputs[c.sht_Terminals])
		self.process_filters(hast_inputs[c.sht_Filters])

		# Process study case details
		self.sc_names = self.get_study_cases(hast_inputs[c.sht_Scenarios])
		self.cont_names = self.get_contingencies(hast_inputs[c.sht_Contingencies])

	def study_settings(self, list_study_settings=None, df_settings=None):
		"""
			Populate study settings
		:param list list_study_settings:
		:param pd.DataFrame df_settings:  DataFrame of study settings for processing
		:return None:
		"""
		# Since this is settings, convert DataFrame to list and extract based on position
		if df_settings is not None:
			l = df_settings[1].tolist()
		else:
			l = list_study_settings

		# Folder to store logs (progress/error) and the excel results if empty will use current working directory
		if not l[0]:
			self.pth_results_folder = os.getcwd() + "\\"
		else:
			self.pth_results_folder = l[0]

		# Leading names to use for exported excel result file (python adds on the unique time and date).
		self.results_name = '{}{}{}.'.format(self.pth_results_folder, l[1], self.uid)
		self.progress_log_name = '{}{}{}.txt'.format(self.pth_results_folder, l[2], self.uid)
		self.error_log_name = '{}{}{}.txt'.format(self.pth_results_folder, l[3], self.uid)
		self.debug_log_name = '{}{}{}.txt'.format(self.pth_results_folder, constants.DEBUG, self.uid)

		# Temporary folder to use to store results exported during script run
		self.pth_results_folder_temp = os.path.join(self.pth_results_folder, self.uid)

		# Constants for power factory
		self.pf_netelm = l[4]
		self.pf_mutelm = '{}{}'.format(l[5], self.uid)
		self.pf_resfolder = '{}{}'.format(l[6], self.uid)
		self.pf_opscen_folder = '{}{}'.format(l[7], self.uid)

		# Constants to control study running
		self.pre_case_check = l[8]
		self.fs_sim = l[9]
		self.hrm_sim = l[10]
		self.skip_failed_lf = l[11]
		self.del_created_folders = l[12]
		self.export_to_excel = l[13]
		self.excel_visible = l[14]
		self.include_rx = l[15]
		self.include_convex_hull = l[16]
		self.export_z = l[17]
		self.export_z12 = l[18]
		self.export_hrm = l[19]

		return None

	def load_flow_settings(self, list_lf_settings):
		"""
			Populate load flow settings
		:param list list_lf_settings:
		:return None:
		"""
		# If there is no value provided then assume
		if not list_lf_settings[0]:
			self.lf.populate_data(load_flow_settings=list_lf_settings[1:])
			self.pf_loadflow_command = None
		else:
			# Settings file for existing load flow settings in PowerFactory
			self.pf_loadflow_command = '{}.{}'.format(list_lf_settings[0], constants.PowerFactory.ldf_command)
			self.lf = None
		return None

	def process_terminals(self, list_of_terminals):
		"""
			Processes the terminals from the HAST input into a dictionary so can lookup the name to use based on
			substation and terminal
		:param list list_of_terminals: List of terminals from HAST inputs, expected in the form
			[name, substation, terminal, include mutual]
		:return None
		"""
		# Get handle for logger
		logger = logging.getLogger(constants.logger_name)
		self.list_of_terms = [TerminalDetails(k[0], k[1], k[2], k[3]) for k in list_of_terminals]
		self.dict_of_terms = {(k.substation, k.terminal): k.name for k in self.list_of_terms}

		# Confirm that none of the terminal names are greater than the maximum allowed character length
		terminal_names = [k.name for k in self.list_of_terms]
		long_names = [x for x in terminal_names if len(x) > constants.HASTInputs.max_terminal_name_length]
		if long_names:

			logger.critical('The following terminal names are greater than the maximum allowed length of {} characters'
							.format(constants.HASTInputs.max_terminal_name_length))
			for x in long_names:
				logger.critical('Terminal name: {}'.format(x))
			raise ValueError(('The terminal names in the HAST inputs {} are too long! Reduce them to less than {} '
							 'characters.').format(self.filename, constants.HASTInputs.max_terminal_name_length))

		# Check all terminal names are unique
		# Get duplicated terminals and report to user then exit
		duplicates = [x for n, x in enumerate(terminal_names) if x in terminal_names[:n]]
		if duplicates:
			msg = ('The user defined Terminal names given in the HAST Inputs workbook {} are not unique for '
				  'each entry.  Please check rename some of the terminals').format(self.filename)
			# Get duplicated entries
			logger.critical(msg)
			logger.critical('The following terminal names have been duplicated:')
			for name in duplicates:
				logger.critical('\t - User Defined Terminal Name: {}'.format(name))
			raise ValueError(msg)

		return None

	def process_filters(self, list_of_filters):
		"""
			Processes the filters from the HAST input into a list of all filters
		:param list list_of_filters: List of handles to type file_io.FilterDetails
		:return None
		"""
		# Get handle for logger
		logger = logging.getLogger(constants.logger_name)
		# Filters already converted to the correct type on initial import so just reference list
		# TODO: Move processing of filters to here rather than initial import
		self.list_of_filters = list_of_filters

		# Check no filter names are duplicated
		filter_names = [k.name for k in self.list_of_filters]
		# Check all filter names are unique
		# Duplicated filter names
		duplicates = [x for n,x in enumerate(filter_names) if x not in filter_names[:n]]
		if duplicates:
			msg = ('The user defined Filter names given in the HAST Inputs workbook {} are not unique for '
				  'each entry.  Please check rename some of the terminals').format(self.filename)
			logger.critical(msg)
			logger.critical('The following names are duplicated:')
			for name in duplicates:
				logger.critical('\t - User Defined Filter Name: {}'.format(name))
			raise ValueError(msg)
		return None

	def vars_to_export(self):
		"""
			Determines the variables that will be exported from PowerFactory and they will be exported in this order
		:return list pf_vars:  Returns list of variables in the format they are defined in PowerFactory
		"""
		c = constants.PowerFactory
		pf_vars = []

		# The order variables are added here determines the order they appear in the export
		# If self impedance data should be exported
		if self.export_z:
			# Whether to include RX data as well
			if self.include_rx:
				pf_vars.append(c.pf_r1)
				pf_vars.append(c.pf_x1)
			pf_vars.append(c.pf_z1)

		# If mutual impedance data should be exported
		if self.export_z12:
			# If RX data should be exported
			if self.include_rx:
				pf_vars.append(c.pf_r12)
				pf_vars.append(c.pf_x12)
			pf_vars.append(c.pf_z12)

		return pf_vars

	def get_study_cases(self, list_of_studycases):
		"""
			Populates dictionary which references all the relevant HAST study case details and then returns a list
			of the names of all the StudyCases that have been considered.
		:return list sc_details:  Returns list of study case names and there corresponding technical details
		"""
		# Get handle for logger
		logger = logging.getLogger(constants.logger_name)

		# If has already been populated then just return the list
		if not self.sc_details:
			# Loop through each row of the imported data
			sc_names = list()
			for sc in list_of_studycases:
				# Transfer row of inputs to class <StudyCaseDetails>
				new_sc = StudyCaseDetails(sc)
				sc_names.append(new_sc.name)
				# Add to dictionary
				self.sc_details[new_sc.name] = new_sc

			# Get list of study_case names and confirm they are all unique
			# Get duplicated study case names
			duplicates = [x for n,x in enumerate(sc_names) if x in sc_names[:n]]
			if duplicates:
				msg = ('The user defined Study Case names given in the HAST Inputs workbook {} are not unique for '
					   'each entry.  Please check rename some of the user defined names').format(self.filename)
				logger.critical(msg)
				logger.critical('The following SC names have been duplicated:')
				for name in duplicates:
					logger.critical('\t - Study Case Name: {}'.format(name))
				raise ValueError(msg)

		return list(self.sc_details.keys())

	def get_contingencies(self, list_of_contingencies):
		"""
			Populates dictionary which references all the relevant HAST study case details and then returns a list
			of the names of all the StudyCases that have been considered.
		:return list sc_details:  Returns list of study case names and there corresponding technical details
		"""
		# Get handle for logger
		logger = logging.getLogger(constants.logger_name)

		# If has already been populated then just return the list
		if not self.cont_details:
			# Loop through each row of the imported data
			cont_names = list()
			for sc in list_of_contingencies:
				# Transfer row of inputs to class <StudyCaseDetails>
				new_cont = ContingencyDetails(sc)
				cont_names.append(new_cont.name)
				# Add to dictionary
				self.cont_details[new_cont.name] = new_cont

			# Get list of contingency names and confirm they are all unique
			# Get duplicated contingency names
			duplicates = [x for n,x in enumerate(cont_names) if x in cont_names[:n]]
			if duplicates:
				msg = ('The user defined Contingency names given in the HAST Inputs workbook {} are not unique for '
					   'each entry.  Please check and rename some of the user defined names').format(self.filename)
				logger.critical(msg)
				logger.critical('The following names that have been provided are duplicated:')
				for name in duplicates:
					logger.critical('\t - Contingency Name: {}'.format(name))
				raise ValueError(msg)


		return list(self.cont_details.keys())


class StudySettings:
	"""
		Class contains the processing of each of the DataFrame items passed as part of the
	"""
	def __init__(self, sht=constants.HASTInputs.study_settings, wkbk=None, pth_file=None):
		# Constants used as part of this
		self.export_folder = str()
		self.results_name = str()
		self.pf_network_elm = str()
		self.pre_case_check = bool()
		self.delete_created_folders = bool()
		self.export_to_excel = bool()
		self.export_rx = bool()
		self.export_mutual = bool()

		self.c = constants.StudySettings
		self.logger = logging.getLogger(constants.logger_name)

		# Unique identifier created from the filename
		self.uid = time.strftime('%y%m%d_%H%M%S')

		# Sheet name
		self.sht = sht

		# Import workbook as dataframe
		if wkbk is None:
			if pth_file:
				wkbk = pd.ExcelFile(pth_file)
				self.pth = pth_file
			else:
				raise IOError('No workbook or path to file provided')
		else:
			# Get workbook path in case path has not been provided
			self.pth = wkbk.io

		# Import Study settings into a DataFrame and process
		self.df = pd.read_excel(
			wkbk, sheet_name=self.sht, index_col=0, usecols=(0, 1), skiprows=4, header=None, squeeze=True
		)

	def process_inputs(self):
		""" Process all of the inputs into attributes of the class """
		# Process results_folder
		self.export_folder = self.process_export_folder()
		self.results_name = self.process_result_name()
		self.pf_network_elm = self.process_net_elements()

		self.pre_case_check = self.process_booleans(key=self.c.pre_case_check)
		self.delete_created_folders = self.process_booleans(key=self.c.delete_created_folders)
		self.export_to_excel = self.process_booleans(key=self.c.export_to_excel)
		self.export_rx = self.process_booleans(key=self.c.export_rx)
		self.export_mutual = self.process_booleans(key=self.c.export_mutual)

		# Sanity check for Boolean values
		self.boolean_sanity_check()

		return None

	def process_export_folder(self, def_value=os.path.join(os.path.dirname(__file__), '..')):
		"""
			Process the export folder and if a blank value is provided then use the default value
		:param str def_value:  Default path to use
		:return str folder:
		"""
		# Get folder from DataFrame, if empty or invalid path then use default folder
		folder = self.df.loc[self.c.export_folder]
		# Normalise path
		def_value = os.path.normpath(def_value)

		if not folder:
			# If no folder provided then use default value
			folder = os.path.normpath(def_value)
		elif not os.path.isdir(os.path.dirname(folder)):
			# If folder provided but not a valid directory for display warning message to user and use default directory
			self.logger.warning((
				'The parent directory for the results export path: {} does not exist and therefore the default directory '
				'of {} has been used instead'
			).format(os.path.dirname(folder), def_value))
			folder = def_value

		# Check if target directory exists and if not create
		if not os.path.isdir(folder):
			os.mkdir(folder)
			self.logger.info('The directory for the results {} does not exist but has been created'.format(folder))

		return folder

	def process_result_name(self, def_value=constants.StudySettings.def_results_name):
		"""
			Processes the results file name
		:param str def_value:  (optional) Default value to use
		:return str results_name:
		"""
		results_name = self.df.loc[self.c.results_name]

		if not results_name:
			# If no value provided then use default value
			self.logger.warning((
				'No value provided in the Input Settings for the results name and so the default value of {} will be '
				'used instead'
			).format(def_value)
			)
			results_name = def_value

		# Add study_time to end of results name
		results_name = '{}_{}{}'.format(results_name, self.uid, constants.Extensions.excel)

		return results_name

	def process_net_elements(self):
		"""
			Processes the details of the folder that contains all the network elements with the appropriate extension
		:return str net_elements:
		"""
		network_folder = str(self.df.loc[self.c.pf_network_elm])

		# TODO: Check if there is an alternative way to handle these network element folders
		if network_folder == '':
			raise ValueError(
				'No value has been provided for the network element folder and it is therefore not possible to identify '
				'the relevant components in PowerFactory'
			)

		if not network_folder.endswith(constants.PowerFactory.pf_network_elements):
			network_folder = '{}.{}'.format(network_folder, constants.PowerFactory.pf_network_elements)

		return network_folder

	def process_booleans(self, key):
		"""
			Function imports the relevant boolean value and confirms it is either True / False, if empty then just
			raises warning message to the user
		:param str key:
		:return bool value:
		"""
		# Get folder from DataFrame, if empty or invalid path then use default folder
		value = self.df.loc[key]

		if value == '':
			value = False
			self.logger.warning(
				(
					'No value has been provided for key item <{}> in worksheet <{}> of the input file <{}> and so '
					'{} will be assumed'
				).format(
					key, self.sht, self.pth, value
				)
			)
		else:
			# Ensure value is a suitable Boolean
			value = bool(value)

		return value

	def boolean_sanity_check(self):
		"""
			Function to check if any of the input Booleans have not been set which require something to be set for
			results to be of any use.

			Throws an error if nothing suitable / logs a warning message

		:return None:
		"""

		if not self.export_to_excel and self.delete_created_folders:
			self.logger.critical((
				'Your input settings for {} = {} means that no results are exported.  However, you are also deleting '
				'all the created results with the command {} = {}.  This is probably not what you meant to happen so '
				'I am stopping here for you to correct your inputs!').format(
				self.c.export_to_excel, self.export_to_excel, self.c.delete_created_folders, self.delete_created_folders)
			)
			raise ValueError('Your input settings mean nothing would be produced')

		if not self.pre_case_check:
			self.logger.warning((
				'You have opted not to run a pre-case check with the input {} = {}, this means that if there are any '
				'issues with an individual contingency the entire studyset may fail'
			).format(self.c.pre_case_check, self.pre_case_check))

		return None


class StudyInputsDev:
	"""
		Class used to import the Settings from the Input Spreadsheet and convert into a format usable elsewhere
	"""
	def __init__(self, pth_file=None):
		"""
			Initialises the settings based on the Study Settings spreadsheet
		:param str file_name:  Path to input settings file
		"""
		# General constants
		self.pth = pth_file
		self.filename = os.path.basename(pth_file)

		with pd.ExcelFile(io=self.pth) as wkbk:
			# Import StudySettings
			self.settings = StudySettings(wkbk=wkbk)
			# TODO: Need to write importers for rest of StudySettings

class StudyCaseDetails:
	def __init__(self, list_of_parameters):
		"""
			Single row of study case parameters imported from spreadsheet are defined into a class for easy lookup
		:param list list_of_parameters:  Single row of inputs
		"""
		self.name = list_of_parameters[0]
		self.project = list_of_parameters[1]
		self.study_case = list_of_parameters[2]
		self.scenario = list_of_parameters[3]

class ContingencyDetails:
	def __init__(self, list_of_parameters):
		"""
			Single row of contingency parameters imported from spreadsheet are defined into a class for easy lookup
		:param list list_of_parameters:  Single row of inputs
		"""
		self.name = list_of_parameters[0]
		self.couplers = []
		for substation, breaker, status in zip(*[iter(list_of_parameters[1:])]*3):
			if substation != '':
				new_coupler = CouplerDetails(substation, breaker, status)
				self.couplers.append(new_coupler)

class CouplerDetails:
	def __init__(self, substation, breaker, status):
		self.substation = substation
		self.breaker = breaker
		self.status = status

class TerminalDetails:
	"""
		Details for each terminal that data is required for from HAST processing
	"""
	def __init__(self, name, substation, terminal, include_mutual=True):
		"""
		:param str name:  HAST Input name to use
		:param str substation:  Name of substation within which terminal is contained
		:param str terminal:   Name of terminal in substation
		:param bool include_mutual:  (optional=True) - If mutual impedance data is not required for this terminal then
			set to False
		"""
		self.name = name
		self.substation = substation
		self.terminal = terminal
		self.include_mutual = include_mutual
		# Reference to PowerFactory established as part of HAST_V2_1.check_terminals
		self.pf_handle = None

class FilterDetails:
	"""
		Class for each filter from the HAST import spreadsheet with a new entry for each substation
	"""
	def __init__(self, row_data):
		"""
			Function to read in the filters and save to list
		:param list row_data:  List of values in the form:
			[name to use for filters,
			substation filter belongs to,
			terminal at which filter should be connected,
			type of filter to use (integer based on PF type),
			Q start, Q stop, number of sizes
			freq start, freq stop, number of freq steps,
			quality factor to use,
			parallel resistance (Rp) value to use
			]
		:return list combined_entry:
		"""
		# Variable initialisation
		self.include = True
		self.nom_voltage = 0.0

		# Confirm row data exists
		if row_data[0] is None:
			self.include = False
			return

		# Name to use for filter
		self.name = row_data[0]
		# Substation and terminal within substation that filter should be connected to
		self.sub ='{}.{}'.format(row_data[1], constants.PowerFactory.pf_substation)
		self.term = '{}.{}'.format(row_data[2], constants.PowerFactory.pf_terminal)
		# Type of filter to use
		self.type = constants.PowerFactory.Filter_type[row_data[3]]
		# Q values for filters (start, stop, no. steps)
		self.q_range = list(np.linspace(row_data[4], row_data[5], row_data[6]))
		self.f_range = list(np.linspace(row_data[7], row_data[8], row_data[9]))
		# Quality factor and parallel resistance values to use
		self.quality_factor = row_data[10]
		self.resistance_parallel = row_data[11]

		# Produce lists of each Q step for each frequency so multiple filters can be tested
		self.f_q_values = list(itertools.product(self.f_range, self.q_range))

class LFSettings:
	def __init__(self):
		"""
			Initialise variables
		"""
		# Target busbar reference found during runtime
		self.rembar = str()

		# Basic
		self.iopt_net = int()  # Calculation method (0 Balanced AC, 1 Unbalanced AC, DC)
		self.iopt_at = int()  # Automatic Tap Adjustment
		self.iopt_asht = int()  # Automatic Shunt Adjustment

		# Added in Automatic Tapping of PSTs but for backwards compatibility will ensure can work when less than 1
		self.iPST_at = int()  # Automatic Tap Adjustment of Phase Shifters

		self.iopt_lim = int()  # Consider Reactive Power Limits
		self.iopt_limScale = int()  # Consider Reactive Power Limits Scaling Factor
		self.iopt_tem = int()  # Temperature Dependency: Line Cable Resistances (0 ...at 20C, 1 at Maximum Operational Temperature)
		self.iopt_pq = int()  # Consider Voltage Dependency of Loads
		self.iopt_fls = int()  # Feeder Load Scaling
		self.iopt_sim = int()  # Consider Coincidence of Low-Voltage Loads
		self.scPnight = int()  # Scaling Factor for Night Storage Heaters

		# Active Power Control
		self.iopt_apdist = int()  # Active Power Control (0 as Dispatched, 1 According to Secondary Control,
		# 2 According to Primary Control, 3 According to Inertia)
		self.iopt_plim = int()  # Consider Active Power Limits
		self.iPbalancing = int()  # (0 Ref Machine, 1 Load, Static Gen, Dist slack by loads, Dist slack by Sync,

		# Get DataObject handle for reference busbar
		self.substation = str()
		self.terminal = str()

		self.phiini = int() # Angle

		# Advanced Options
		self.i_power = int()  # Load Flow Method ( NR Current, 1 NR (Power Eqn Classic)
		self.iopt_notopo = int()  # No Topology Rebuild
		self.iopt_noinit = int()  # No initialisation
		self.utr_init = int()  # Consideration of transformer winding ratio
		self.maxPhaseShift = int()  # Max Transformer Phase Shift
		self.itapopt = int()  # Tap Adjustment ( 0 Direct, 1 Step)
		self.krelax = int()  # Min Controller Relaxation Factor

		self.iopt_stamode = int()  # Station Controller (0 Standard, 1 Gen HV, 2 Gen LV
		self.iopt_igntow = int()  # Modelling Method of Towers (0 With In/ Output signals, 1 ignore couplings, 2 equation in lines)
		self.initOPF = int()  # Use this load flow for initialisation of OPF
		self.zoneScale = int()  # Zone Scaling ( 0 Consider all loads, 1 Consider adjustable loads only)

		# Iteration Control
		self.itrlx = int()  # Max No Iterations for Newton-Raphson Iteration
		self.ictrlx = int()  # Max No Iterations for Outer Loop
		self.nsteps = int()  # Max No Iterations for Number of steps

		self.errlf = int()  # Max Acceptable Load Flow Error for Nodes
		self.erreq = int()  # Max Acceptable Load Flow Error for Model Equations
		self.iStepAdapt = int()  # Iteration Step Size ( 0 Automatic, 1 Fixed Relaxation)
		self.relax = int()  # If Fixed Relaxation factor
		self.iopt_lev = int()  # Automatic Model Adaptation for Convergence

		# Outputs
		self.iShowOutLoopMsg = int()  # Show 'outer Loop' Messages
		self.iopt_show = int()  # Show Convergence Progress Report
		self.num_conv = int()  # Number of reported buses/models per iteration
		self.iopt_check = int()  # Show verification report
		self.loadmax = int()  # Max Loading of Edge Element
		self.vlmin = int()  # Lower Limit of Allowed Voltage
		self.vlmax = int()  # Upper Limit of Allowed Voltage
		self.iopt_chctr = int()  # Check Control Conditions

		# Load Generation Scaling
		self.scLoadFac = int()  # Load Scaling Factor
		self.scGenFac = int()  # Generation Scaling Factor
		self.scMotFac = int()  # Motor Scaling Factor

		# Low Voltage Analysis
		self.Sfix = int()  # Fixed Load kVA
		self.cosfix = int()  # Power Factor of Fixed Load
		self.Svar = int()  # Max Power Per Customer kVA
		self.cosvar = int()  # Power Factor of Variable Part
		self.ginf = int()  # Coincidence Factor
		self.i_volt = int()  # Voltage Drop Analysis (0 Stochastic Evaluation, 1 Maximum Current Estimation)

		# Advanced Simulation Options
		self.iopt_prot = int()  # Consider Protection Devices ( 0 None, 1 all, 2 Main, 3 Backup)
		self.ign_comp = int()  # Ignore Composite Elements

	def populate_data(self, load_flow_settings):
		"""
			List of settings for the load flow from HAST if using a manual settings file
		:param list load_flow_settings:
		"""
		# Loadflow settings
		# -------------------------------------------------------------------------------------
		# Create new object for the load flow on the base case so that existing settings are not overwritten

		# Get handle for load flow command from study case
		# Basic
		self.iopt_net = load_flow_settings[0]  # Calculation method (0 Balanced AC, 1 Unbalanced AC, DC)
		self.iopt_at = load_flow_settings[1]  # Automatic Tap Adjustment
		self.iopt_asht = load_flow_settings[2]  # Automatic Shunt Adjustment

		# Added in Automatic Tapping of PSTs with default values
		self.iPST_at = load_flow_settings[3]  # Automatic Tap Adjustment of Phase Shifters

		self.iopt_lim = load_flow_settings[4]  # Consider Reactive Power Limits
		self.iopt_limScale = load_flow_settings[5]  # Consider Reactive Power Limits Scaling Factor
		self.iopt_tem = load_flow_settings[6]  # Temperature Dependency: Line Cable Resistances (0 ...at 20C, 1 at Maximum Operational Temperature)
		self.iopt_pq = load_flow_settings[7]  # Consider Voltage Dependency of Loads
		self.iopt_fls = load_flow_settings[8]  # Feeder Load Scaling
		self.iopt_sim = load_flow_settings[9]  # Consider Coincidence of Low-Voltage Loads
		self.scPnight = load_flow_settings[10]  # Scaling Factor for Night Storage Heaters

		# Active Power Control
		self.iopt_apdist = load_flow_settings[11]  # Active Power Control (0 as Dispatched, 1 According to Secondary Control,
		# 2 According to Primary Control, 3 According Inertia)
		self.iopt_plim = load_flow_settings[12]  # Consider Active Power Limits
		self.iPbalancing = load_flow_settings[13]  # (0 Ref Machine, 1 Load, Static Gen, Dist slack by loads, Dist slack by Sync,

		# Get DataObject handle for reference busbar
		net_folder_name, substation, terminal = load_flow_settings[14].split('\\')
		# Confirm that substation and terminal types exist in name
		if not substation.endswith(constants.PowerFactory.pf_substation):
			self.substation = '{}.{}'.format(substation, constants.PowerFactory.pf_substation)
		if not terminal.endswith(constants.PowerFactory.pf_terminal):
			self.terminal = '{}.{}'.format(terminal, constants.PowerFactory.pf_terminal)


		self.phiini = load_flow_settings[15]  # Angle

		# Advanced Options
		self.i_power = load_flow_settings[16]  # Load Flow Method ( NR Current, 1 NR (Power Eqn Classic)
		self.iopt_notopo = load_flow_settings[17]  # No Topology Rebuild
		self.iopt_noinit = load_flow_settings[18]  # No initialisation
		self.utr_init = load_flow_settings[19]  # Consideration of transformer winding ratio
		self.maxPhaseShift = load_flow_settings[20]  # Max Transformer Phase Shift
		self.itapopt = load_flow_settings[21]  # Tap Adjustment ( 0 Direct, 1 Step)
		self.krelax = load_flow_settings[22]  # Min Controller Relaxation Factor

		self.iopt_stamode = load_flow_settings[23]  # Station Controller (0 Standard, 1 Gen HV, 2 Gen LV
		self.iopt_igntow = load_flow_settings[24]  # Modelling Method of Towers (0 With In/ Output signals, 1 ignore couplings, 2 equation in lines)
		self.initOPF = load_flow_settings[25]  # Use this load flow for initialisation of OPF
		self.zoneScale = load_flow_settings[26]  # Zone Scaling ( 0 Consider all loads, 1 Consider adjustable loads only)

		# Iteration Control
		self.itrlx = load_flow_settings[27]  # Max No Iterations for Newton-Raphson Iteration
		self.ictrlx = load_flow_settings[28]  # Max No Iterations for Outer Loop
		self.nsteps = load_flow_settings[29]  # Max No Iterations for Number of steps

		self.errlf = load_flow_settings[30]  # Max Acceptable Load Flow Error for Nodes
		self.erreq = load_flow_settings[31]  # Max Acceptable Load Flow Error for Model Equations
		self.iStepAdapt = load_flow_settings[32]  # Iteration Step Size ( 0 Automatic, 1 Fixed Relaxation)
		self.relax = load_flow_settings[33]  # If Fixed Relaxation factor
		self.iopt_lev = load_flow_settings[34]  # Automatic Model Adaptation for Convergence

		# Outputs
		self.iShowOutLoopMsg = load_flow_settings[35]  # Show 'outer Loop' Messages
		self.iopt_show = load_flow_settings[36]  # Show Convergence Progress Report
		self.num_conv = load_flow_settings[37]  # Number of reported buses/models per iteration
		self.iopt_check = load_flow_settings[38]  # Show verification report
		self.loadmax = load_flow_settings[39]  # Max Loading of Edge Element
		self.vlmin = load_flow_settings[40]  # Lower Limit of Allowed Voltage
		self.vlmax = load_flow_settings[41]  # Upper Limit of Allowed Voltage
		# ldf.outcmd =  load_flow_settings[42-offset]          		# Output
		self.iopt_chctr = load_flow_settings[43]  # Check Control Conditions
		# ldf.chkcmd = load_flow_settings[44-offset]            	# Command

		# Load Generation Scaling
		self.scLoadFac = load_flow_settings[45]  # Load Scaling Factor
		self.scGenFac = load_flow_settings[46]  # Generation Scaling Factor
		self.scMotFac = load_flow_settings[47]  # Motor Scaling Factor

		# Low Voltage Analysis
		self.Sfix = load_flow_settings[48]  # Fixed Load kVA
		self.cosfix = load_flow_settings[49]  # Power Factor of Fixed Load
		self.Svar = load_flow_settings[50]  # Max Power Per Customer kVA
		self.cosvar = load_flow_settings[51]  # Power Factor of Variable Part
		self.ginf = load_flow_settings[52]  # Coincidence Factor
		self.i_volt = load_flow_settings[53]  # Voltage Drop Analysis (0 Stochastic Evaluation, 1 Maximum Current Estimation)

		# Advanced Simulation Options
		self.iopt_prot = load_flow_settings[54]  # Consider Protection Devices ( 0 None, 1 all, 2 Main, 3 Backup)
		self.ign_comp = load_flow_settings[55]  # Ignore Composite Elements


	def find_reference_terminal(self, app):
		"""
			Find and populate reference terminal for machine
		:param powerfactory.app app:
		:return None:
		"""
		pf_sub = app.GetCalcRelevantObjects(self.substation)
		pf_term = pf_sub[0].GetContents(self.terminal)[0]
		self.rembar = pf_term

		return None

#  ----- UNIT TESTS -----
class TestExcelSetup(unittest.TestCase):
	"""
		UnitTest to test the operation of various excel workbook functions
	"""

	def test_hast_settings_import(self):
		"""
			Tests that excel will import setting appropriately
		"""
		pth = os.path.dirname(os.path.abspath(__file__))
		pth_test_files = 'tests'
		test_workbook = 'HAST_Inputs.xlsx'
		input_file = os.path.join(pth, pth_test_files, test_workbook)

		xl = Excel(print_info=print, print_error=print)
		analysis_dict = xl.import_excel_harmonic_inputs(pth_workbook=input_file)
		self.assertEqual(len(analysis_dict.keys()), 8)
