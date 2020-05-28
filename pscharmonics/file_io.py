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

import os
import numpy as np
import itertools
import pscharmonics.constants as constants
import logging
import pandas as pd
import collections


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

def update_duplicates(key, df):
	"""
		Function will look for any duplicates in a particular column and then append a number to everything after
		the first one
	:param str key:  Column to lookup
	:param pd.DataFrame df: DataFrame to be processed
	:return pd.DataFrame, bool df_updated, updated: Updated DataFrame and status flag to show updated for log messages
	"""
	# Empty list and initialised value to show no changes
	dfs = list()
	updated = False

	# Group data frame by key value
	# TODO: Do we need to ensure the order of the overall DataFrame remains unchanged
	for _, df_group in df.groupby(key):
		# Extract all key values into list and then loop through to append new value
		names = list(df_group.loc[:, key])
		if len(names) > 1:
			updated = True
			# Check if number of entries is greater than 1, i.e. duplicated
			for i in range(1, len(names)):
				names[i] = '{}({})'.format(names[i], i)

		# Produce new DataFrame with updated names and add to list
		# (copy statement to avoid updating original dataframe during loop)
		df_updated = df_group.copy()
		df_updated.loc[:, key] = names
		dfs.append(df_updated)

	# Combine returned DataFrames
	df_updated = pd.concat(dfs)
	# Ensure order remains as originally
	df_updated.sort_index(axis=0, inplace=True)

	return df_updated, updated


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

# class StudyInputs:
# 	"""
# 		Class used to import the Settings from the Input Spreadsheet and convert into a format usable elsewhere
# 	"""
# 	def __init__(self, hast_inputs=None, uid_time=constants.uid, filename=''):
# 		"""
# 			Initialises the settings based on the HAST Study Settings spreadsheet
# 		:param dict hast_inputs:  Dictionary of input data returned from file_io.Excel.import_excel_harmonic_inputs
# 		:param str uid_time:  Time string to use as the uid for these files
# 		:param str filename:  Filename of the HAST Inputs file used from which this data is extracted
# 		"""
# 		c = constants.PowerFactory
# 		# General constants
# 		self.filename=filename
#
# 		self.uid = uid_time
#
# 		# Attribute definitions (study settings)
# 		self.pth_results_folder = str()
# 		self.results_name = str()
# 		self.progress_log_name = str()
# 		self.error_log_name = str()
# 		self.debug_log_name = str()
# 		self.pth_results_folder_temp = str()
# 		self.pf_netelm = str()
# 		self.pf_mutelm = str()
# 		self.pf_resfolder = str()
# 		self.pf_opscen_folder = str()
# 		self.pre_case_check = bool()
# 		self.fs_sim = bool()
# 		self.hrm_sim = bool()
# 		self.skip_failed_lf = bool()
# 		self.del_created_folders = bool()
# 		self.export_to_excel = bool()
# 		self.excel_visible = bool()
# 		self.include_rx = bool()
# 		self.include_convex_hull = bool()
# 		self.export_z = bool()
# 		self.export_z12 = bool()
# 		self.export_hrm = bool()
#
# 		# Attribute definitions (study_case_details)
# 		self.sc_details = dict()
# 		self.sc_names = list()
#
# 		# Attribute definitions (contingency_details)
# 		self.cont_details = dict()
# 		self.cont_names = list()
#
# 		# Attribute definitions (terminals)
# 		self.list_of_terms = list()
# 		self.dict_of_terms = dict()
#
# 		# Attribute definitions (filters)
# 		self.list_of_filters = list()
#
# 		# Load Flow Settings
# 		# Will contain full string to load flow command to be used
# 		self.pf_loadflow_command = str()
# 		# Will contain reference to LFSettings which contains all settings
# 		self.lf = LFSettings()
#
# 		# Process study settings
# 		self.study_settings(hast_inputs[c.sht_Study])
#
# 		# Process load flow settings
# 		self.load_flow_settings(hast_inputs[c.sht_LF])
#
# 		# Process List of Terminals
# 		self.process_terminals(hast_inputs[c.sht_Terminals])
# 		self.process_filters(hast_inputs[c.sht_Filters])
#
# 		# Process study case details
# 		self.sc_names = self.get_study_cases(hast_inputs[c.sht_Scenarios])
# 		self.cont_names = self.get_contingencies(hast_inputs[c.sht_Contingencies])
#
# 	def study_settings(self, list_study_settings=None, df_settings=None):
# 		"""
# 			Populate study settings
# 		:param list list_study_settings:
# 		:param pd.DataFrame df_settings:  DataFrame of study settings for processing
# 		:return None:
# 		"""
# 		# Since this is settings, convert DataFrame to list and extract based on position
# 		if df_settings is not None:
# 			l = df_settings[1].tolist()
# 		else:
# 			l = list_study_settings
#
# 		# Folder to store logs (progress/error) and the excel results if empty will use current working directory
# 		if not l[0]:
# 			self.pth_results_folder = os.getcwd() + "\\"
# 		else:
# 			self.pth_results_folder = l[0]
#
# 		# Leading names to use for exported excel result file (python adds on the unique time and date).
# 		self.results_name = '{}{}{}.'.format(self.pth_results_folder, l[1], self.uid)
# 		self.progress_log_name = '{}{}{}.txt'.format(self.pth_results_folder, l[2], self.uid)
# 		self.error_log_name = '{}{}{}.txt'.format(self.pth_results_folder, l[3], self.uid)
# 		self.debug_log_name = '{}{}{}.txt'.format(self.pth_results_folder, constants.DEBUG, self.uid)
#
# 		# Temporary folder to use to store results exported during script run
# 		self.pth_results_folder_temp = os.path.join(self.pth_results_folder, self.uid)
#
# 		# Constants for power factory
# 		self.pf_netelm = l[4]
# 		self.pf_mutelm = '{}{}'.format(l[5], self.uid)
# 		self.pf_resfolder = '{}{}'.format(l[6], self.uid)
# 		self.pf_opscen_folder = '{}{}'.format(l[7], self.uid)
#
# 		# Constants to control study running
# 		self.pre_case_check = l[8]
# 		self.fs_sim = l[9]
# 		self.hrm_sim = l[10]
# 		self.skip_failed_lf = l[11]
# 		self.del_created_folders = l[12]
# 		self.export_to_excel = l[13]
# 		self.excel_visible = l[14]
# 		self.include_rx = l[15]
# 		self.include_convex_hull = l[16]
# 		self.export_z = l[17]
# 		self.export_z12 = l[18]
# 		self.export_hrm = l[19]
#
# 		return None
#
# 	def load_flow_settings(self, list_lf_settings):
# 		"""
# 			Populate load flow settings
# 		:param list list_lf_settings:
# 		:return None:
# 		"""
# 		# If there is no value provided then assume
# 		if not list_lf_settings[0]:
# 			self.lf.populate_data(load_flow_settings=list_lf_settings[1:])
# 			self.pf_loadflow_command = None
# 		else:
# 			# Settings file for existing load flow settings in PowerFactory
# 			self.pf_loadflow_command = '{}.{}'.format(list_lf_settings[0], constants.PowerFactory.ldf_command)
# 			self.lf = None
# 		return None
#
# 	def process_terminals(self, list_of_terminals):
# 		"""
# 			Processes the terminals from the HAST input into a dictionary so can lookup the name to use based on
# 			substation and terminal
# 		:param list list_of_terminals: List of terminals from HAST inputs, expected in the form
# 			[name, substation, terminal, include mutual]
# 		:return None
# 		"""
# 		# Get handle for logger
# 		logger = logging.getLogger(constants.logger_name)
# 		self.list_of_terms = [TerminalDetails(k[0], k[1], k[2], k[3]) for k in list_of_terminals]
# 		self.dict_of_terms = {(k.substation, k.terminal): k.name for k in self.list_of_terms}
#
# 		# Confirm that none of the terminal names are greater than the maximum allowed character length
# 		terminal_names = [k.name for k in self.list_of_terms]
# 		long_names = [x for x in terminal_names if len(x) > constants.HASTInputs.max_terminal_name_length]
# 		if long_names:
#
# 			logger.critical('The following terminal names are greater than the maximum allowed length of {} characters'
# 							.format(constants.HASTInputs.max_terminal_name_length))
# 			for x in long_names:
# 				logger.critical('Terminal name: {}'.format(x))
# 			raise ValueError(('The terminal names in the HAST inputs {} are too long! Reduce them to less than {} '
# 							 'characters.').format(self.filename, constants.HASTInputs.max_terminal_name_length))
#
# 		# Check all terminal names are unique
# 		# Get duplicated terminals and report to user then exit
# 		duplicates = [x for n, x in enumerate(terminal_names) if x in terminal_names[:n]]
# 		if duplicates:
# 			msg = ('The user defined Terminal names given in the HAST Inputs workbook {} are not unique for '
# 				  'each entry.  Please check rename some of the terminals').format(self.filename)
# 			# Get duplicated entries
# 			logger.critical(msg)
# 			logger.critical('The following terminal names have been duplicated:')
# 			for name in duplicates:
# 				logger.critical('\t - User Defined Terminal Name: {}'.format(name))
# 			raise ValueError(msg)
#
# 		return None
#
# 	def process_filters(self, list_of_filters):
# 		"""
# 			Processes the filters from the HAST input into a list of all filters
# 		:param list list_of_filters: List of handles to type file_io.FilterDetails
# 		:return None
# 		"""
# 		# Get handle for logger
# 		logger = logging.getLogger(constants.logger_name)
# 		# Filters already converted to the correct type on initial import so just reference list
# 		# TODO: Move processing of filters to here rather than initial import
# 		self.list_of_filters = list_of_filters
#
# 		# Check no filter names are duplicated
# 		filter_names = [k.name for k in self.list_of_filters]
# 		# Check all filter names are unique
# 		# Duplicated filter names
# 		duplicates = [x for n,x in enumerate(filter_names) if x not in filter_names[:n]]
# 		if duplicates:
# 			msg = ('The user defined Filter names given in the HAST Inputs workbook {} are not unique for '
# 				  'each entry.  Please check rename some of the terminals').format(self.filename)
# 			logger.critical(msg)
# 			logger.critical('The following names are duplicated:')
# 			for name in duplicates:
# 				logger.critical('\t - User Defined Filter Name: {}'.format(name))
# 			raise ValueError(msg)
# 		return None
#
# 	def vars_to_export(self):
# 		"""
# 			Determines the variables that will be exported from PowerFactory and they will be exported in this order
# 		:return list pf_vars:  Returns list of variables in the format they are defined in PowerFactory
# 		"""
# 		c = constants.PowerFactory
# 		pf_vars = []
#
# 		# The order variables are added here determines the order they appear in the export
# 		# If self impedance data should be exported
# 		if self.export_z:
# 			# Whether to include RX data as well
# 			if self.include_rx:
# 				pf_vars.append(c.pf_r1)
# 				pf_vars.append(c.pf_x1)
# 			pf_vars.append(c.pf_z1)
#
# 		# If mutual impedance data should be exported
# 		if self.export_z12:
# 			# If RX data should be exported
# 			if self.include_rx:
# 				pf_vars.append(c.pf_r12)
# 				pf_vars.append(c.pf_x12)
# 			pf_vars.append(c.pf_z12)
#
# 		return pf_vars
#
# 	def get_study_cases(self, list_of_studycases):
# 		"""
# 			Populates dictionary which references all the relevant HAST study case details and then returns a list
# 			of the names of all the StudyCases that have been considered.
# 		:return list sc_details:  Returns list of study case names and there corresponding technical details
# 		"""
# 		# Get handle for logger
# 		logger = logging.getLogger(constants.logger_name)
#
# 		# If has already been populated then just return the list
# 		if not self.sc_details:
# 			# Loop through each row of the imported data
# 			sc_names = list()
# 			for sc in list_of_studycases:
# 				# Transfer row of inputs to class <StudyCaseDetails>
# 				new_sc = StudyCaseDetails(sc)
# 				sc_names.append(new_sc.name)
# 				# Add to dictionary
# 				self.sc_details[new_sc.name] = new_sc
#
# 			# Get list of study_case names and confirm they are all unique
# 			# Get duplicated study case names
# 			duplicates = [x for n,x in enumerate(sc_names) if x in sc_names[:n]]
# 			if duplicates:
# 				msg = ('The user defined Study Case names given in the HAST Inputs workbook {} are not unique for '
# 					   'each entry.  Please check rename some of the user defined names').format(self.filename)
# 				logger.critical(msg)
# 				logger.critical('The following SC names have been duplicated:')
# 				for name in duplicates:
# 					logger.critical('\t - Study Case Name: {}'.format(name))
# 				raise ValueError(msg)
#
# 		return list(self.sc_details.keys())
#
# 	def get_contingencies(self, list_of_contingencies):
# 		"""
# 			Populates dictionary which references all the relevant HAST study case details and then returns a list
# 			of the names of all the StudyCases that have been considered.
# 		:return list sc_details:  Returns list of study case names and there corresponding technical details
# 		"""
# 		# Get handle for logger
# 		logger = logging.getLogger(constants.logger_name)
#
# 		# If has already been populated then just return the list
# 		if not self.cont_details:
# 			# Loop through each row of the imported data
# 			cont_names = list()
# 			for sc in list_of_contingencies:
# 				# Transfer row of inputs to class <StudyCaseDetails>
# 				new_cont = ContingencyDetails(sc)
# 				cont_names.append(new_cont.name)
# 				# Add to dictionary
# 				self.cont_details[new_cont.name] = new_cont
#
# 			# Get list of contingency names and confirm they are all unique
# 			# Get duplicated contingency names
# 			duplicates = [x for n,x in enumerate(cont_names) if x in cont_names[:n]]
# 			if duplicates:
# 				msg = ('The user defined Contingency names given in the HAST Inputs workbook {} are not unique for '
# 					   'each entry.  Please check and rename some of the user defined names').format(self.filename)
# 				logger.critical(msg)
# 				logger.critical('The following names that have been provided are duplicated:')
# 				for name in duplicates:
# 					logger.critical('\t - Contingency Name: {}'.format(name))
# 				raise ValueError(msg)
#
#
# 		return list(self.cont_details.keys())


class StudySettings:
	"""
		Class contains the processing of each of the DataFrame items passed as part of the
	"""
	def __init__(self, sht=constants.HASTInputs.study_settings, wkbk=None, pth_file=None, gui_mode=False):
		"""
			Process the worksheet to extract the relevant StudyCase details
		:param str sht:  (optional) Name of worksheet to use
		:param pd.ExcelFile wkbk:  (optional) Handle to workbook
		:param str pth_file: (optional) Handle to workbook
		:param bool gui_mode: (optional) When set to True this will prevent the creation of a new
							folder since the user will provide the folder when running the studies
		"""
		# Constants used as part of this
		self.export_folder = str()
		self.results_name = str()
		self.pre_case_check = bool()
		self.delete_created_folders = bool()
		self.export_to_excel = bool()
		self.export_rx = bool()
		self.export_mutual = bool()
		self.include_intact = bool()

		self.c = constants.StudySettings
		self.logger = logging.getLogger(constants.logger_name)

		# Sheet name
		self.sht = sht

		# GUI Mode - Prevents creation of folders that are defined later
		self.gui_mode = gui_mode

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
		self.process_inputs()

	def process_inputs(self):
		""" Process all of the inputs into attributes of the class """
		# Process results_folder
		if not self.gui_mode:
			self.export_folder = self.process_export_folder()
		else:
			self.logger.debug('Running in GUI mode and therefore export_folder is defined later')
		# self.results_name = self.process_result_name()
		# self.pf_network_elm = self.process_net_elements()

		self.pre_case_check = self.process_booleans(key=self.c.pre_case_check)
		self.delete_created_folders = self.process_booleans(key=self.c.delete_created_folders)
		self.export_to_excel = self.process_booleans(key=self.c.export_to_excel)
		self.export_rx = self.process_booleans(key=self.c.export_rx)
		self.export_mutual = self.process_booleans(key=self.c.export_mutual)
		self.include_intact = self.process_booleans(key=self.c.include_intact)

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
		def_value = os.path.join(os.path.normpath(def_value), constants.uid)

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

	# def process_result_name(self, def_value=constants.StudySettings.def_results_name):
	# 	"""
	# 		Processes the results file name
	# 	:param str def_value:  (optional) Default value to use
	# 	:return str results_name:
	# 	"""
	# 	results_name = self.df.loc[self.c.results_name]
	#
	# 	if not results_name:
	# 		# If no value provided then use default value
	# 		self.logger.warning((
	# 			'No value provided in the Input Settings for the results name and so the default value of {} will be '
	# 			'used instead'
	# 		).format(def_value)
	# 		)
	# 		results_name = def_value
	#
	# 	# Add study_time to end of results name
	# 	results_name = '{}_{}{}'.format(results_name, self.uid, constants.Extensions.excel)
	#
	# 	return results_name

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
				'issues with an individual contingency the entire study set may fail'
			).format(self.c.pre_case_check, self.pre_case_check))

		return None

class StudyInputsDev:
	"""
		Class used to import the Settings from the Input Spreadsheet and convert into a format usable elsewhere
	"""
	def __init__(self, pth_file=None, gui_mode=False):
		"""
			Initialises the settings based on the Study Settings spreadsheet
		:param str pth_file:  Path to input settings file
		:param bool gui_mode: (optional) when set to True this will creating the export folder since that
							will be processed later
		"""
		# General constants
		self.pth = pth_file
		self.filename = os.path.basename(pth_file)
		self.logger = logging.getLogger(constants.logger_name)

		# TODO: Adjust to update if an Import Error occurs
		self.error = False

		self.logger.info('Importing settings from file: {}'.format(self.pth))

		with pd.ExcelFile(io=self.pth) as wkbk:
			# Import StudySettings
			self.settings = StudySettings(wkbk=wkbk, gui_mode=gui_mode)
			self.cases = self.process_study_cases(wkbk=wkbk)
			self.contingency_cmd, self.contingencies = self.process_contingencies(wkbk=wkbk)
			self.terminals = self.process_terminals(wkbk=wkbk)
			self.lf_settings = self.process_lf_settings(wkbk=wkbk)
			self.fs_settings = self.process_fs_settings(wkbk=wkbk)

		if not self.error:
			self.logger.info('Importing settings from file: {} completed'.format(self.pth))
		else:
			self.logger.warning('Error during import of settings from file: {}'.format(self.pth))

	def load_workbook(self, pth_file=None):
		"""
			Function to load the workbook and return a handle to it
		:param str pth_file: (optional) File path to workbook
		:return pd.ExcelFile wkbk: Handle to workbook
		"""
		if pth_file:
			# Load the workbook using pd.ExcelFile and return the reference to the workbook
			wkbk = pd.ExcelFile(pth_file)
			self.pth = pth_file
		else:
			raise IOError('No workbook or path to file provided')

		return wkbk

	def process_study_cases(self, sht=constants.HASTInputs.study_cases, wkbk=None, pth_file=None):
		"""
			Function imports the DataFrame of study cases and then separates each one into it's own study case.  These
			are then returned as a dictionary with the name being used as the key.

			These inputs are based on the Scenarios detailed in the Inputs spreadsheet

			:param str sht:  (optional) Name of worksheet to use
			:param pd.ExcelFile wkbk:  (optional) Handle to workbook
			:param str pth_file: (optional) Handle to workbook
		:return pd.DataFrame df:  Returns a DataFrame of the study cases so can group by columns
		"""

		# Import workbook as dataframe
		if wkbk is None:
			wkbk = self.load_workbook(pth_file=pth_file)

		# Import Study settings into a DataFrame and process, do not need to worry about unique columns since done by
		# position
		df = pd.read_excel(wkbk, sheet_name=sht, skiprows=3, header=0, usecols=(0, 1, 2, 3))
		cols = constants.StudySettings.studycase_columns
		df.columns = cols

		# Process df names to confirm unique
		name_key = cols[0]
		df, updated = update_duplicates(key=name_key, df=df)

		# Make index of study_cases name after duplicates removed
		df.set_index(constants.StudySettings.name, inplace=True, drop=False)

		if updated:
			self.logger.warning(
				(
					'Duplicated names for StudyCases provided in the column {} of worksheet <{}> and so some have been '
					'renamed so the new list of names is:\n\t{}'
				).format(name_key, sht, '\n\t'.join(df.loc[:, name_key].values))
			)

		# # Iterate through each DataFrame and create a study case instance and OrderedDict used to ensure no change in
		# # order
		# study_cases = collections.OrderedDict()
		# for key, item in df.iterrows():
		# 	new_case = StudyCaseDetails(list_of_parameters=item.values)
		# 	study_cases[new_case.name] = new_case

		return df

	def process_contingencies(self, sht=constants.HASTInputs.contingencies, wkbk=None, pth_file=None):
		"""
			Function imports the DataFrame of contingencies and then each into its own contingency under a single
			dictionary key.  Each dictionary item contains the relevant outages to be taken in the form
			ContingencyDetails.

			Also returns a string detailing the name of a Contingency file if one is provided

		:param str sht:  (optional) Name of worksheet to use
		:param pd.ExcelFile wkbk:  (optional) Handle to workbook
		:param str pth_file: (optional) File path to workbook
		:return (str, dict) (contingency_cmd, contingencies):  Returns both:
																The command for all contingencies
																A dictionary with each of the outages
		"""

		# Import workbook as dataframe
		if wkbk is None:
			wkbk = self.load_workbook(pth_file=pth_file)

		# Get details of existing contingency command and if exists set the reference command
		# Squeeze command to ensure converted to a series since only a single row is being imported
		df = pd.read_excel(
			wkbk, sheet_name=sht, skiprows=2, nrows=1, usecols=(0,1,2,3), index_col=None, header=None
		).squeeze(axis=0)
		cmd = df.iloc[3]
		# Valid command is a string whereas non-valid commands will be either False or an empty string
		if cmd:
			contingency_cmd = cmd
		else:
			contingency_cmd = str()


		# Import rest of details of contingencies into a DataFrame and process, do not need to worry about unique
		# columns since done by position
		df = pd.read_excel(wkbk, sheet_name=sht, skiprows=3, header=(0, 1))
		cols = df.columns

		# Process df names to confirm unique
		name_key = cols[0]
		df, updated = update_duplicates(key=name_key, df=df)

		# Make index of contingencies name after duplicates removed
		df.set_index(name_key, inplace=True, drop=False)


		if updated:
			self.logger.warning(
				(
					'Duplicated names for Contingencies provided in the column {} of worksheet <{}> and so some have been '
					'renamed and the new list of names is:\n\t{}'
				).format(name_key, sht, '\n\t'.join(df.loc[:, name_key].values))
			)

		# Iterate through each DataFrame and create a study case instance and OrderedDict used to ensure no change in order
		contingencies = collections.OrderedDict()
		for key, item in df.iterrows():
			cont = ContingencyDetails(list_of_parameters=item.values)
			contingencies[cont.name] = cont

		return contingency_cmd, contingencies

	def process_terminals(self, sht=constants.HASTInputs.terminals, wkbk=None, pth_file=None):
		"""
			Function imports the DataFrame of terminals.
			These are then returned as a dictionary with the name being used as the key.

			These inputs are based on the Scenarios detailed in the Inputs spreadsheet

		:param str sht:  (optional) Name of worksheet to use
		:param pd.ExcelFile wkbk:  (optional) Handle to workbook
		:param str pth_file: (optional) File path to workbook
		:return dict study_cases:
		"""

		# Import workbook as dataframe
		if wkbk is None:
			wkbk = self.load_workbook(pth_file=pth_file)

		# Import Contingencies into a DataFrame and process, do not need to worry about unique columns since done by
		# position
		df = pd.read_excel(wkbk, sheet_name=sht, skiprows=3, header=0)
		cols = df.columns

		# Process df names to confirm unique
		name_key = cols[0]
		df, updated = update_duplicates(key=name_key, df=df)

		# Make index of terminals name after duplicates removed
		df.set_index(name_key, inplace=True, drop=False)

		if updated:
			self.logger.warning(
				(
					'Duplicated names for Terminals provided in the column {} of worksheet <{}> and so some have been '
					'renamed and the new list of names is:\n\t{}'
				).format(name_key, sht, '\n\t'.join(df.loc[:, name_key].values))
			)

		# Iterate through each DataFrame and create a study case instance and OrderedDict used to ensure no change in order
		terminals = collections.OrderedDict()
		for key, item in df.iterrows():
			term = TerminalDetails(list_of_parameters=item.values)
			terminals[term.name] = term

		return terminals

	def process_lf_settings(self, sht=constants.HASTInputs.lf_settings, wkbk=None, pth_file=None):
		"""
			Process the provided load flow settings
		:param str sht:  (optional) Name of worksheet to use
		:param pd.ExcelFile wkbk:  (optional) Handle to workbook
		:param str pth_file: (optional) File path to workbook
		:return LFSettings lf_settings:  Returns reference to instance of load_flow settings
		"""
		# Import workbook as dataframe
		if wkbk is None:
			wkbk = self.load_workbook(pth_file=pth_file)

		# Import Load flow settings into a DataFrame and process
		df = pd.read_excel(
			wkbk, sheet_name=sht, usecols=(3, ), skiprows=3, header=None, squeeze=True
		)

		lf_settings = LFSettings(existing_command=df.iloc[0], detailed_settings=df.iloc[1:])
		return lf_settings

	def process_fs_settings(self, sht=constants.HASTInputs.fs_settings, wkbk=None, pth_file=None):
		"""
			Process the provided frequency sweep settings
		:param str sht:  (optional) Name of worksheet to use
		:param pd.ExcelFile wkbk:  (optional) Handle to workbook
		:param str pth_file: (optional) File path to workbook
		:return LFSettings lf_settings:  Returns reference to instance of load_flow settings
		"""
		# Import workbook as dataframe
		if wkbk is None:
			wkbk = self.load_workbook(pth_file=pth_file)

		# Import Load flow settings into a DataFrame and process
		df = pd.read_excel(
			wkbk, sheet_name=sht, usecols=(3, ), skiprows=3, header=None, squeeze=True
		)

		settings = FSSettings(existing_command=df.iloc[0], detailed_settings=df.iloc[1:])
		return settings

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
		# Status flag set to True if cannot be found, contingency fails, etc.
		self.not_included = False

		self.name = list_of_parameters[0]
		self.couplers = []
		for substation, breaker, status in zip(*[iter(list_of_parameters[1:])]*3):
			if substation != '' and breaker != '' and str(breaker) != 'nan':
				new_coupler = CouplerDetails(substation, breaker, status)
				self.couplers.append(new_coupler)

		# If contingency has been defined then needs to be included in results
		# # Check if this contingency relates to the intact system in which case it will be skipped
		if len(self.couplers) == 0 or self.name == constants.HASTInputs.base_case:
			self.skip = True
		else:
			self.skip = False

class CouplerDetails:
	def __init__(self, substation, breaker, status):
		"""
			Define the substations, breaker and status
		:param str substation:  Name of substation
		:param str breaker: Name of breaker
		:param bool status: Status breaker is changed to
		"""
		# Check if substation already has substation type ending and if not add it
		if not str(substation).endswith(constants.PowerFactory.pf_substation):
			substation = '{}.{}'.format(substation, constants.PowerFactory.pf_substation)

		# Check if breaker ends in the correct format
		if not str(breaker).endswith(constants.PowerFactory.pf_coupler):
			breaker = '{}.{}'.format(breaker, constants.PowerFactory.pf_coupler)

		# Confirm that status for operation is either true or false
		if status == constants.HASTInputs.switch_open:
			status = False
		elif status == constants.HASTInputs.switch_close:
			status = True
		else:
			logger = logging.getLogger(constants.logger_name)
			logger.warning(
				(
					'The breaker <{}> associated with substation <{}> has a value of {} which is not expected value of '
					'{} or {}.  The operation is assumed to be {} for this study but the user should check '
					'that is what they intended'
				).format(
					breaker, substation, status,
					constants.HASTInputs.switch_open, constants.HASTInputs.switch_close,
					constants.HASTInputs.switch_open
				)
			)
			status = False

		self.substation = substation
		self.breaker = breaker
		self.status = status

class TerminalDetails:
	"""
		Details for each terminal that data is required for from HAST processing
	"""
	def __init__(self, name=str(), substation=str(), terminal=str(), include_mutual=True, list_of_parameters=list()):
		"""
			Process each terminal
		:param list list_of_parameters: (optional=none)
		:param str name:  HAST Input name to use
		:param str substation:  Name of substation within which terminal is contained
		:param str terminal:   Name of terminal in substation
		:param bool include_mutual:  (optional=True) - If mutual impedance data is not required for this terminal then
			set to False
		"""

		if len(list_of_parameters) > 0:
			name = str(list_of_parameters[0])
			substation = str(list_of_parameters[1])
			terminal = str(list_of_parameters[2])
			include_mutual = bool(list_of_parameters[3])

		# Add in the relevant endings if they do not already exist
		c = constants.PowerFactory
		if not substation.endswith(c.pf_substation):
			substation = '{}.{}'.format(substation, c.pf_substation)

		if not terminal.endswith(c.pf_terminal):
			terminal = '{}.{}'.format(terminal, c.pf_terminal)

		self.name = name
		self.substation = substation
		self.terminal = terminal
		self.include_mutual = include_mutual

		# The following are populated once looked for in the specific PowerFactory project
		self.pf_handle = None
		self.found = None

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
	def __init__(self, existing_command, detailed_settings):
		"""
			Initialise variables
		:param str existing_command:  Reference to an existing command where it already exists
		:param list detailed_settings:  Settings to be used where existing command does not exist
		"""
		self.logger = logging.getLogger(constants.logger_name)

		# Add the Load Flow command to string
		if existing_command:
			if not existing_command.endswith('.{}'.format(constants.PowerFactory.ldf_command)):
				existing_command = '{}.{}'.format(existing_command, constants.PowerFactory.ldf_command)
		else:
			existing_command = None
		self.cmd = existing_command

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

		# Value set to True if error occurs when processing settings
		self.settings_error = False

		try:
			self.populate_data(load_flow_settings=list(detailed_settings))
		except (IndexError, AttributeError):
			if self.cmd:
				self.logger.warning(
					(
						'There were some missing settings in the load flow settings input but an existing command {} '
						'is being used instead. If this command is missing from the Study Cases then the script will '
						'fail'
					).format(self.cmd)
				)
				self.settings_error = True
			else:
				self.logger.warning(
					(
						'The load flow settings provided are incorrect / missing some values and no input for an existing '
						'load flow command (.{}) has been provided.  The study will just use the default load flow '
						'command associated with each study case but results may be inconsistent!'
					).format(constants.PowerFactory.ldf_command)
				)
				self.settings_error = True


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
		ref_machine = load_flow_settings[14]

		# To avoid error when passed empty string or empty cell which is imported as pd.na
		if not ref_machine or pd.isna(ref_machine):
			self.substation = None
			self.terminal = None
		else:
			net_folder_name, substation, terminal = ref_machine.split('\\')
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
		self.iopt_chctr = load_flow_settings[42]  # Check Control Conditions

		# Load Generation Scaling
		self.scLoadFac = load_flow_settings[43]  # Load Scaling Factor
		self.scGenFac = load_flow_settings[44]  # Generation Scaling Factor
		self.scMotFac = load_flow_settings[45]  # Motor Scaling Factor

		# Low Voltage Analysis
		self.Sfix = load_flow_settings[46]  # Fixed Load kVA
		self.cosfix = load_flow_settings[47]  # Power Factor of Fixed Load
		self.Svar = load_flow_settings[48]  # Max Power Per Customer kVA
		self.cosvar = load_flow_settings[49]  # Power Factor of Variable Part
		self.ginf = load_flow_settings[50]  # Coincidence Factor
		self.i_volt = load_flow_settings[51]  # Voltage Drop Analysis (0 Stochastic Evaluation, 1 Maximum Current Estimation)

		# Advanced Simulation Options
		self.iopt_prot = load_flow_settings[52]  # Consider Protection Devices ( 0 None, 1 all, 2 Main, 3 Backup)
		self.ign_comp = load_flow_settings[53]  # Ignore Composite Elements

	def find_reference_terminal(self, app):
		"""
			Find and populate reference terminal for machine
		:param powerfactory.app app:
		:return None:
		"""
		# Confirm that a machine has actually been provided
		if self.substation is None or self.terminal is None:
			pf_term = None
			self.logger.debug('No reference machine provided in inputs')

		else:
			pf_sub = app.GetCalcRelevantObjects(self.substation)

			if len(pf_sub) == 0:
				pf_term = None
			else:
				pf_term = pf_sub[0].GetContents(self.terminal)
				if len(pf_term) == 0:
					pf_term = None

			if pf_term is None:
				self.logger.warning(
					(
						'Either the substation {} or terminal {} detailed as being the reference busbar cannot be '
						'found and therefore no reference busbar will be defined in the PowerFactory load flow'
					).format(self.substation, self.terminal)
				)

		return pf_term

class FSSettings:
	""" Class contains the frequency scan settings """
	def __init__(self, existing_command, detailed_settings):
		"""
			Initialise variables
		:param str existing_command:  Reference to an existing command where it already exists
		:param list detailed_settings:  Settings to be used where existing command does not exist
		"""
		self.logger = logging.getLogger(constants.logger_name)

		# Add the Load Flow command to string
		if existing_command:
			if not existing_command.endswith('.{}'.format(constants.PowerFactory.fs_command)):
				existing_command = '{}.{}'.format(existing_command, constants.PowerFactory.fs_command)
		else:
			existing_command = None
		self.cmd = existing_command

		# Nominal frequency of system
		self.frnom = float()

		# Basic
		self.iopt_net = int() # Network representation (0 Balanced, 1 Unbalanced)
		self.fstart = float() # Impedance calculation start frequency
		self.fstop = float() # Impedance calculation stop frequency
		self.fstep = float() # Impedance calculation step size
		self.i_adapt = bool() # Automatic step size adaption (False = No, True = Yes)

		# Advanced
		self.errmax = float() # Setting for step size adaption maximum prediction error
		self.errinc = float() # Setting for step size minimum prediction error
		self.ninc = int() # Step size increase delay
		self.ioutall = bool() # Calculate R, X at output frequency for all nodes (False = No, True = Yes)

		# Value set to True if error occurs when processing settings
		self.settings_error = False

		try:
			self.populate_data(fs_settings=list(detailed_settings))
		except (IndexError, AttributeError):
			if self.cmd:
				self.logger.warning(
					(
						'There were some missing settings in the frequency sweep settings input but an existing command {} '
						'is being used instead. If this command is missing from the Study Cases then the script will '
						'fail'
					).format(self.cmd)
				)
				self.settings_error = True
			else:
				self.logger.warning(
					(
						'The frequency sweep settings provided are incorrect / missing some values and no input for an '
						'existing frequency sweep command (.{}) has been provided.  The study will just use the default '
						'frequency sweep command associated with each study case but results may be inconsistent!'
					).format(constants.PowerFactory.fs_command)
				)
				self.settings_error = True


	def populate_data(self, fs_settings):
		"""
			List of settings for the load flow from HAST if using a manual settings file
		:param list fs_settings:
		"""

		# Nominal frequency of system
		self.frnom = fs_settings[0]

		# Basic
		self.iopt_net = int(fs_settings[1])
		self.fstart = fs_settings[2]
		self.fstop = fs_settings[3]
		self.fstep = fs_settings[4]
		self.i_adapt = bool(fs_settings[5])

		# Advanced
		self.errmax = fs_settings[6]
		self.errinc = fs_settings[7]
		self.ninc = int(fs_settings[8])
		self.ioutall = bool(fs_settings[9])

class ResultsExport:
	"""
		Class to deal with handling of the results that are created
	"""
	def __init__(self, pth, name):
		"""
			Initialises, if path passed then results saved there
		:param str pth:  Path to where the results should be saved
		:param str name:  Name that should be used for the results file
		"""
		self.logger = logging.getLogger(constants.logger_name)

		self.parent_pth = pth
		self.results_name = name

		self.results_folder = str()

	def create_results_folder(self):
		"""
			Routine creates the sub_folder which will contain all of the detailed results
		:return None:
		"""

		self.results_folder = os.path.join(self.parent_pth, self.results_name)

		# Check if already exists and if not create it
		if not os.path.exists(self.results_folder):
			os.mkdir(self.results_folder)

		return None

