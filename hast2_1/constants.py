"""
#######################################################################################################################
###											Constants																###
###		Central point to store all constants associated with HAST													###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###		project JI6973 for EirGrid project PSPF010 - Specialise Support in Power Quality Analysis during 2019		###
###																													###
#######################################################################################################################
"""

import pandas as pd
import numpy as np
import os
import unittest

# Meta Data
__author__ = 'David Mills'
__email__ = 'david.mills@pscconsulting.com'
__phone__ = '+44 7899 984158'
__status__ = 'Constants'

nom_freq = 50.0
logger_name = 'HAST'

# When parallel processing will ensure this number of cpus are kept free
cpu_keep_free = 1
# If unable to calculate this is the assumed maximum number of processes
default_max_processes = 3

class PowerFactory:
	"""
		Constants used in this script
	"""
	# String used to define the tuning frequency of the filter
	hz='Hz'
	sht_Filters = 'Filters'
	sht_Terminals = 'Terminals'
	sht_Scenarios = 'Base_Scenarios'
	sht_Contingencies = 'Contingencies'
	sht_Study = 'Study_Settings'
	sht_LF = 'Loadflow_Settings'
	sht_Freq = 'Frequency_Sweep'
	sht_HLF = 'Harmonic_Loadflow'
	HAST_Input_Scenario_Sheets = (sht_Contingencies, sht_Scenarios, sht_Terminals, sht_Filters)
	HAST_Input_Settings_Sheets = (sht_Study, sht_LF, sht_Freq, sht_HLF)
	# Different filter types available in PowerFactory 2016
	Filter_type = {'C-Type':4,
				   'Single':0,
				   'High Pass':3}
	pf_substation = 'ElmSubstat'
	pf_terminal =  'ElmTerm'
	pf_coupler = 'ElmCoup'
	pf_mutual = 'ElmMut'
	pf_case = 'IntCase'
	pf_scenario = 'IntScenario'
	pf_filter = 'ElmShnt'
	pf_cubicle = 'StaCubic'
	pf_term_voltage = 'uknom'
	pf_shn_term = 'bus1'
	pf_shn_voltage = 'ushnm'
	pf_shn_type = 'shtype'
	pf_shn_q = 'qtotn'
	pf_shn_freq = 'fres'
	pf_shn_tuning = 'nres'
	pf_shn_qfactor = 'greaf0'
	pf_shn_qfactor_nom = 'grea'
	pf_shn_rp = 'rpara'
	# constants for variations
	pf_scheme = 'IntScheme'
	pf_stage = 'IntSstage'
	pf_results = 'ElmRes'

	# General Types
	pf_folder_type = 'IntFolder'
	pf_prjfolder_type = 'netdat'

	# Default results file name
	default_results_name = 'HAST_Res'
	default_fs_extension = '_FS'
	default_hldf_extension = '_HLDF'

	pf_r1 = 'm:R'
	pf_x1 = 'm:X'
	pf_z1 = 'm:Z'
	pf_r12 = 'c:R_12'
	pf_x12 = 'c:X_12'
	pf_z12 = 'c:Z_12'
	pf_nom_voltage = 'e:uknom'

	class ComRes:
		# Power Factory class name
		pf_comres = 'ComRes'
		#Com Res setting constants

		# File export type:
		#	0 = Output window
		#	1 = Windows clipboard
		#	2 = Measurement file (ElmFile)
		#	3 = Comtrade
		#	4 = Testfile
		#	5 = PSSPLT Version 2.0
		#	6 = Commas Separated Values (*.csv)
		export_type = 'iopt_exp'
		# Name of file to export to (if appropriate)
		file = 'f_name'
		# Type of deparators to use (0 = Custom, 1 = system defaults)
		separators = 'iopt_sep'
		# Export object headers only (0 = all data, 1 = headers only)
		object_head_only = 'iopt_honly'
		# Variables to extract (0 = all, 1 = custom list)
		variables_all = 'iopt_csel'
		# Name of result file from PF to export
		result = 'pResult'
		# Details to export from element:
		# 	0 = None,
		# 	1 = Name,
		# 	2 = Short path and name,
		# 	3 = Path and name,
		# 	4 = Foreign key
		element = 'iopt_locn'
		# Details to export from variable:
		#	0 = None,
		#	1 = Parameters name,
		#	3 = Short description,
		#	4 = Full description
		variable = 'ciopt_head'
		# Custom of full dataset (0 = full, 1 = custom)
		user_interval = 'iopt_tsel'
		# Export values (0 = values, 1 = variable descriptors only)
		export_values = 'iopt_vars'
		# Shift time of results (0 = none, 1 = shift)
		shift_time = 'iopt_rscl'
		# Filter time of results (0 = None, 1 = filter)
		filtered_time = 'filtered'

class ResultsExtract:
	"""
		Constants used in processing the results
	"""
	study_types = ('FS', 'HLF')
	extension = '.xlsx'
	# Labels used for frequency scan results extract
	lbl_StudyCase = 'Study Case'
	lbl_Frequency = 'Frequency in Hz'
	lbl_Filter_ID = 'Filter Details'
	lbl_Contingency = 'Contingency'
	lbl_FullName = 'Full Result Name'
	lbl_Terminal = 'Terminal Name'
	lbl_Reference_Terminal = 'Terminal'
	lbl_Result = 'Result Type'
	idx_nom_voltage = 'Nom Voltage (kV)'
	# Location of m:R, m:X, m:Z, etc.
	loc_pf_variable = 4
	# Location of m:R12, m:X12, m:Z12, etc.
	loc_pf_variable_mutual = loc_pf_variable + 1
	loc_contingency = 1

	# Chart grouping
	chart_grouping = (lbl_StudyCase, lbl_Contingency)
	chart_grouping_base_case = (lbl_Contingency, )

	# Default positions
	start_row = 31 # (0 referenced so will be Excel row 32)
	start_col = 0 # (0 referenced so will be Excel col A)
	col_spacing = 2 # Leaves 1 empty column between results

	# Labels for charts
	chart_type = {'type': 'scatter'}
	lbl_Impedance = 'Impendance in Ohms'

	# Positioning of charts in excel workbook
	chrt_row = 1
	chrt_col = 1
	# Number of columns between each chart
	chrt_space = 18
	chrt_vert_space = 30

	# Labels for processing
	# This label is used for the column headers when an entry should be deleted post processing
	lbl_to_delete = 'TO DELETE'

	# Chart plot properties
	line_width = 1.0
	lbl_position = 'next_to'
	chrt_width = 960
	chrt_height = 576

	# Based on details here:  https://xlsxwriter.readthedocs.io/chart.html
	grid_lines = {'visible': True,
				  'line': {
					  'width': 0.75,
					  'dash_type': 'dash'}
				  }

	def __init__(self):
		"""
			Initial class
		"""
		self.color_map = dict()

	def get_color_map(self, pth_color_map=None, refresh=False):
		"""
			Obtains a dictionary of the colors to use for plotting graphs in excel.
		:param: str pth_color_map: (Optional=None) - If a different color map is desired can be passed as input
		:return dict color_map:  Returns a dictionary of the color map based on N-1 contingency : hex color code
		"""
		def hex_converter(value):
			"""
				Used to convert the number to a hex value including leading # during import of excel
			:param str value:  Value to be converted
			:return str value:  Returns value with leading #
			"""
			# Confirm haven't got a nan value before returning so that nan values can be removed
			if value is np.nan:
				return value
			else:
				return '#{}'.format(value)

		if not refresh and len(self.color_map) > 0:
			return self.color_map

		# If no color map has been provided then use the default one in the script directory
		if pth_color_map is None:
			pth_color_map = os.path.join(os.path.dirname(os.path.realpath(__file__)),'N1_color_map.xlsm')

		# Import data into a DataFrame in case there is any other processing that needs to be done
		df_colormap = pd.read_excel(pth_color_map, header=0,
									usecols=1, converters={1: hex_converter})
		# Set the index of the dataframe equal to the first column
		df_colormap.set_index(df_colormap.columns[0], inplace=True)
		# Remove any nan values so that only actual colous remain and the length of the dataframe can be used
		# to determine the plots
		df_colormap.dropna(axis=0, inplace=True)

		# Extract the index and color values
		index = df_colormap.index
		values = df_colormap.iloc[:,0]

		# Produce dictionary for lookup and return
		self.color_map=dict(zip(index, values))
		return self.color_map

class HASTInputs:
	file_name = 'HAST Inputs'
	file_format = '.xlsx'
	base_case = 'Base_Case'
	mutual_variables = ["c:Z_12", "c:R_12", "c:X_12"]
	fs_term_variables = ["m:R", "m:X", "m:Z", "m:phiz", "e:uknom"]
	hldf_term_variables = ['m:HD', 'm:THD']
	res_values = ['b:fnow','b:ifnow']
	# For checking variable extraction only
	all_variable_types = (mutual_variables +
						  fs_term_variables +
						  hldf_term_variables +
						  res_values)
	# Names of worksheets which contain the relevant inputs
	terminals = 'Terminals'
	study_settings = 'Study_Settings'
	study_cases = 'Base_Scenarios'
	contingencies = 'Contingencies'
	# Maximum length of an objects name in PowerFactory 2016 is 40 characters.
	# Therefore the maximum name that can be used for a single terminal is 19 characters to allow two terminals to be
	# joined together
	max_terminal_name_length = 19
	# Default value on whether mutual impedance data should be included or not
	default_include_mutual = True


analysis_sheets = (
	(PowerFactory.sht_Study, "B5"),
	(PowerFactory.sht_Scenarios, "A5"),
	(PowerFactory.sht_Contingencies, "A5"),
	(PowerFactory.sht_Terminals, "A5"),
	(PowerFactory.sht_LF, "D5"),
	(PowerFactory.sht_Freq, "D5"),
	(PowerFactory.sht_HLF, "D5"),
	(PowerFactory.sht_Filters, "A5"))

iec_limits = [
	["IEC", "61000-3-6", "Harmonics", "THD", 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18,
	 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40],
	["IEC", "61000-3-6", "Limits", 3, 1.4, 2, 0.8, 2, 0.4, 2, 0.4, 1, 0.35, 1.5, 0.32, 1.5, 0.3, 0.3, 0.28, 1.2,
	 0.265, 0.93, 0.255, 0.2, 0.246, 0.88, 0.24, 0.816, 0.233, 0.2, 0.227, 0.703, 0.223, 0.66, 0.219, 0.2, 0.2158,
	 0.58, 0.2127, 0.55, 0.21, 0.2, 0.2075]]

# Format to use to mark debug log
DEBUG = 'DEBUG'

class TestResultsExtract(unittest.TestCase):
	"""
		Unit Test to test the operation and constant definition of the Results Extract class
	"""
	@classmethod
	def setUpClass(cls):
		"""
			Setup the handle for the Results Extract class
		:return:
		"""
		cls.res_extract = ResultsExtract()

	def test_results_extract_constant(self):
		"""
			Simple test to confirm test case is operational
		"""
		self.assertEqual(self.res_extract.lbl_StudyCase, ResultsExtract.lbl_StudyCase)

	def test_color_map(self):
		"""
			Test to confirm import of N-1 Color map successful and convertion into a
			dictionary of values works correctly
		"""
		dict_color_map = self.res_extract.get_color_map()
		self.assertEqual('#000000', dict_color_map[0])

