"""
#######################################################################################################################
###											Constants																###
###		Central point to store all constants associated with PSC harmonics											###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###																													###
#######################################################################################################################
"""

import pandas as pd
import numpy as np
import os
import sys
import glob
import time

# Label used when displaying messages
__title__ = 'PSC Automated Frequency Scan Tool'
__version__ = '1.3.0'

logger_name = 'PSC'
logger = None

DEBUG = False

# Unique identifier populated for each study run
uid = time.strftime('%Y%m%d_%H%M%S')

# Reference to local directory used by other packages
local_directory=os.path.abspath(os.path.dirname(__file__))

#

class General:
	# Value that is used as leading value
	cmd_lf_leader = 'PSC_LF'
	cmd_fs_leader = 'PSC_FS'
	cmd_cont_leader = 'PSC_Cont'
	cmd_fsres_leader = 'PSC_FS_Res'
	cmd_contres_leader = 'PSC_Cont_Res'
	cmd_autotasks_leader = 'PSC_Auto'

	# Names to use for export columns
	prj = 'Project'
	sc = 'Study Case'
	op = 'Operating Scenario'

	# Default names to use for log messages
	debug_log = 'DEBUG'
	progress_log = 'INFO'
	error_log = 'ERROR'

	user_guide_reference='JA7896-03 PSC Harmonics User Guide.pdf'
	user_guide_folder = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'docs'))
	user_guide_pth = os.path.join(user_guide_folder, user_guide_reference)

	# These are the threshold at which log messages will either be warned about or deleted
	threshold_warning = 500
	threshold_delete = 700
	file_number_thresholds = (threshold_warning, threshold_delete)

	# Nominal frequency assumed for studies
	nominal_frequency = 50.0


class PowerFactory:
	"""
		Constants used in this script
	"""
	# Constants relating to the paths
	pf_year = 2019
	year_max_tested = 2019
	pf_service_pack = ''
	dig_path = str()
	dig_python_path = str()
	# Populated with available installed PowerFactory versions on initialisation
	available_power_factory_versions = list()
	target_power_factory = 'PowerFactory 2019'

	# The following list details python versions which are non compatible
	non_compatible_python_versions = ['3.5']

	# Default PowerFactory installation directories
	default_install_directory = r'C:\Program Files\DIgSILENT'
	power_factory_search = 'PowerFactory 20*'

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
	# Different filter types available in PowerFactory 2016
	Filter_type = {'C-Type':4,
				   'Single':0,
				   'High Pass':3}
	pf_substation = 'ElmSubstat'
	pf_line = 'ElmLne'
	pf_branch = 'ElmBranch'
	pf_terminal =  'ElmTerm'
	pf_coupler = 'ElmCoup'
	pf_mutual = 'ElmMut'
	pf_fault_event = 'IntEvt'
	pf_switch_event = 'EvtSwitch'
	pf_outage_event = 'EvtOutage'
	pf_case = 'IntCase'
	pf_scenario = 'IntScenario'
	pf_filter = 'ElmShnt'
	pf_cubicle = 'StaCubic'
	pf_term_voltage = 'uknom'
	pf_shn_term = 'bus1'
	pf_shn_voltage = 'ushnm'
	pf_shn_type = 'shtype'
	pf_shn_q = 'qtotn'
	pf_shn_inputmode = 'mode_inp'
	pf_shn_selectedinput = 'DES'
	pf_shn_freq = 'fres'
	pf_shn_tuning = 'nres'
	pf_shn_qfactor = 'greaf0'
	pf_shn_qfactor_nom = 'grea'
	pf_shn_rp = 'rpara'
	# constants for variations
	pf_scheme = 'IntScheme'
	pf_stage = 'IntSstage'
	pf_results = 'ElmRes'
	pf_network_elements = 'ElmNet'
	pf_project = 'IntPrj'
	# Command for carrying out contingency analysis and applying each outage
	pf_cont_analysis = 'ComSimoutage'
	pf_outage = 'ComOutage'

	# General Types
	pf_folder_type = 'IntFolder'
	pf_fault_cases_folder = 'IntFltcases'
	pf_netdata_folder_type = 'netdat'
	pf_faults_folder_type = 'fault'
	pf_sc_folder_type = 'study'
	pf_os_folder_type = 'scen'

	# Default results file name
	default_fs_extension = '_FS'

	pf_r1 = 'm:R'
	pf_x1 = 'm:X'
	pf_z1 = 'm:Z'
	pf_r12 = 'c:R_12'
	pf_x12 = 'c:X_12'
	pf_z12 = 'c:Z_12'
	pf_nom_voltage = 'e:uknom'
	pf_freq = 'b:fnow in Hz'
	pf_harm = 'b:ifnow'

	ldf_command = 'ComLdf'
	hldf_command = 'ComHldf'
	fs_command = 'ComFsweep'
	autotasks_command = 'ComTasks'

	# Folder names for temporary folders
	temp_sc_folder = 'temp_sc'
	temp_os_folder = 'temp_os'
	temp_faults_folder = 'temp_faults'
	temp_mutual_folder = 'mutual_elements'

	# Constants associated with the handling of PowerFactory initialisation and
	# potential intermittent errors
	# Number of attempts to obtain a license
	license_activation_attempts = 5
	# Number of seconds to wait between license attempts
	license_activation_delay = 5.0
	# Error codes which could be intermittent and therefore the script should try again
	# Description in PowerFactory help file:  ErrorCodeReference_en.pdf
	license_activation_error_codes = (3000, 3002, 3005, 3011, 3012, 4000, 4002, 5000)

	# Each results variable has a default type and need to assign the defaults to the newly created results
	# variables
	def_results_hlf = 5		# Harmonic load flow
	def_results_fs = 9		# Frequency sweep
	def_results_cont = 13#

	# User default settings
	user_default_settings = 'Set\Def\Settings.SetUser'

	# Number of seconds to allow when waiting for parallel processor response
	parallel_time_out = 100

	# This is a maximum impedance value, above this and it is assumed to be open circuit and will be ignored
	max_impedance = 1E6

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
		# Type of separators to use (0 = Custom, 1 = system defaults)
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

	def __init__(self):
		"""
			Initialises the relevant python paths depending on the version that has been loaded
		"""
		# Get reference to logger
		self.logger = logger

		# Find all PowerFactory versions installed in this location
		power_factory_paths = glob.glob(os.path.join(self.default_install_directory, self.power_factory_search))
		self.available_power_factory_versions = [os.path.basename(x) for x in power_factory_paths]
		self.available_power_factory_versions.sort()

	def select_power_factory_version(self, pf_version=None, mock_python_version=str()):
		"""
			Function allows the user to select a specific PowerFactory version, if none is selected then
			the most recent version of PowerFactory is used
		:param str pf_version: (optional) - If None then the most recent PowerFactory version is used
		:param str mock_python_version:  For TESTING only, gets replaced with a different version to check correct
										errors thrown if incorrect version provided

		"""

		# If no pf_version is provided then the default version defined is used if it exists in the available versions
		# otherwise the latest version
		if pf_version is None:
			if self.target_power_factory not in self.available_power_factory_versions:
				# Rather than assuming a particular version just default to the latest version
				self.target_power_factory = self.available_power_factory_versions[-1]
		elif pf_version in self.available_power_factory_versions:
			self.target_power_factory = pf_version
		else:
			self.logger.critical(
				(
					'The PowerFactory version {} has been selected but does not exist in the installed versions:\n\t{}'
				).format(pf_version, '\n\t'.join([x for x in self.available_power_factory_versions]))
			)
			raise EnvironmentError('Invalid PowerFactory version')

		self.logger.debug('PowerFactory version <{}> will be used'.format(self.target_power_factory))

		# Find year from selected PowerFactory version
		# pf_version is assumed to take the format PowerFactory #### and therefore the #### can be extracted
		year = [int(s) for s in self.target_power_factory.split() if s.isdigit()][0]

		# Confirm the year is > 2017 and < 2020 otherwise warn that hasn't been fully tested
		if int(year) < 2018:
			self.logger.warning(
				(
					'You are using PowerFactory version {}.\n'
					'In the 2018 version there were some API changes which have been considered in this script.  The '
					'previous versions may still work but they have not been considered as part of the development '
					'testing and so you are advised to carefully check your results.'
				).format(year)
			)
		elif int(year) > self.year_max_tested:
			self.logger.warning(
				(
					'You are using PowerFactory version {}.\n'
					'This script has only been tested up to year {} and therefore changes in the PowerFactory API may '
					'impact on the running and results you produce.  You are advised to check the results carefully or '
					'consider updating the developments testing for this version.  For further details contact:\n{}'
				).format(year, self.year_max_tested, Author.contact_summary)
			)

		# Find the installation directory for the PowerFactory paths
		self.dig_path = os.path.join(self.default_install_directory, self.target_power_factory)

		# Now checks for Python versions within this PowerFactory
		if mock_python_version:
			# Running in a test mode to check an error is created
			self.logger.warning('TESTING - Testing with a mock python version to raise exception, if not expected then there is a '
								'script error! - TESTING')
			current_python_version = mock_python_version
		else:
			# Formulate string for python version
			current_python_version = '{}.{}'.format(sys.version_info.major, sys.version_info.minor)


		# Get list of supported python versions
		list_of_available_versions = [os.path.basename(x) for x in glob.glob(os.path.join(self.dig_path, 'Python', '*'))]
		if current_python_version in self.non_compatible_python_versions:
			self.logger.critical(
				(
					'You are using Python version {}, this script is not compatible with that version or the following '
					'versions: \n\t Python {}\n  Additionally, the PowerFactory version you have selected ({}) is only compatible '
					'with the following Python versions: \n\t Python {}'
				).format(
					current_python_version, '\n\t Python '.join(self.non_compatible_python_versions),
					self.target_power_factory, '\n\t Python '.join(list_of_available_versions)
				)
			)
			raise EnvironmentError('Non Compatible Python Version')

		# Define the Python Path for PowerFactory
		self.dig_python_path = os.path.join(self.dig_path, 'Python', current_python_version)
		if not os.path.isdir(self.dig_python_path):

			self.logger.critical(
				(
					'You are running python version: {} but only the following versions are supported by this version of'
					'PowerFactory ({}):\n\t Python {}'
				).format(current_python_version, self.target_power_factory, '\n\t Python '.join(list_of_available_versions))
			)
			raise EnvironmentError('Incompatible Python version')

class Contingencies:
	""" Contains constants associated with naming of contingencies used in export """
	# Name to give for base_case / intact system condition
	intact = 'Intact'
	# This is the default contingency number that is given for the intact study case
	intact_cont_num = -1

	prj = General.prj
	sc = General.sc
	op = General.op
	cont = 'Contingency'
	idx = 'Contingency Number'
	status = 'Convergent'

	# Contingencies are provided for either circuit breakers or lines, these are provided in a dictionary which
	# uses the following keys to identify them
	cb = 'CB'
	lines = 'Lines'

	# Columns that are used for the contingency headers
	df_columns = [
		prj, sc, op, cont, idx, status
	]

	# Maximum number of contingencies before which studies will be run using parallel processing
	parallel_threshold = 50

	# Variables to keep from cont_results
	col_object = 'b:i_obj'
	col_number = 'b:number'
	col_nonconvergent = 'b:inoconv'

	# Name that is used for worksheet when exporting details of convergent contingencies
	export_sheet_name = 'Contingencies'

class Terminals:
	""" Contains constants associated with processing of terminals used in export """
	prj = General.prj
	name = 'Terminal / Mutual Name'
	sub1 = 'Substation 1 Name'
	sub2 = 'Substation 2 Name'
	bus1 = 'Busbar 1 Name'
	bus2 = 'Busbar 2 Name'
	include_mutual = 'Include Mutual'
	status = 'Found'
	planned_name = 'Planned Mutual Impedance Name'

	# Columns that are used for the contingency headers
	columns = [
		name, sub1, bus1, include_mutual, status, planned_name, sub2, bus2
	]

	# Character used to join terminals together
	join_char = '_'

	# In PowerFactory 2016 (and potentially others) there is a max terminal name length of 40 characters and therefore
	# this is the name that is used for the terminal couplings
	max_coupled_length = 40

	# When trimming terminals this is the minimum length their name will be trimmed to
	min_term_length = 4

	# Name that is used for worksheet when exporting details of missing terminals
	export_sheet_name = 'Terminals'

class Results:
	"""
		Constants used in processing the results
	"""
	# Labels for all results details
	skipped = 'Study Skipped'

	study_fs = 'FS'
	# Symbol used to join study_case name with contingency name
	joiner = '_'
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
	lbl_Harmonic_Order = 'Harmonic Order'
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
	row_spacing = 2 # Leaves empty rows between DataFrame extracts

	# Labels for charts
	chart_type = {'type': 'scatter'}
	lbl_Impedance = 'Impedance in Ohms'
	lbl_Resistance = 'Resistance in Ohms'
	lbl_Reactance = 'Reactance in Ohms'

	# Positioning of charts in excel workbook
	chrt_row = 1
	chrt_col = 1
	# Number of columns between each chart
	chrt_space = 18
	chrt_vert_space = 30

	# Number of plots to include vertically before resetting counter
	loci_chrt_space = 9
	loci_plots_vertically = 4

	# Labels for processing
	# This label is used for the column headers when an entry should be deleted post processing
	lbl_to_delete = 'TO DELETE'

	# Chart plot properties
	line_width = 1.0
	lbl_position = 'next_to'
	chrt_width = 960
	chrt_height = 576

	# Dimensions for loci plots
	chrt_loci_width = 576
	chrt_loci_height = 576

	# Based on details here:  https://xlsxwriter.readthedocs.io/chart.html
	grid_lines = {'visible': True,
				  'line': {
					  'width': 0.75,
					  'dash_type': 'dash'}
				  }

	# Marker settings for R and X values
	marker_type = 'x'
	market_size = 5

	# Font size for chart title
	font_size_chart_title = 14



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
		# Remove any nan values so that only actual colours remain and the length of the dataframe can be used
		# to determine the plots
		df_colormap.dropna(axis=0, inplace=True)

		# Extract the index and color values
		index = df_colormap.index
		values = df_colormap.iloc[:,0]

		# Produce dictionary for lookup and return
		self.color_map=dict(zip(index, values))
		return self.color_map

class StudyInputs:
	file_name = 'Inputs'
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
	# Contingencies can be specified by either identifying specific breakers or lines
	cont_breakers = 'Contingencies_Breakers'
	cont_lines = 'Contingencies_Lines'
	lf_settings = 'Loadflow_Settings'
	fs_settings = 'Frequency_Sweep'
	loci_settings = 'Loci_Settings'
	# Maximum length of an objects name in PowerFactory 2016 is 40 characters.
	# Therefore the maximum name that can be used for a single terminal is 19 characters to allow two terminals to be
	# joined together
	max_terminal_name_length = 19
	# Default value on whether mutual impedance data should be included or not
	default_include_mutual = True

	# Default value for automatic tap changing of PST
	def_automatic_pst_tap = 1

	# Text use to define an open or close status for a switch
	switch_open = 'Open'
	switch_close = 'Close'

	# Text used to define whether a line is in service or out of service
	in_service = 'In Service'
	out_of_service = 'Out of Service'

class LociInputs:
	# Default values to use for impedance loci processing
	# Defaults to using +/- half of nominal frequency
	def_polygon_range = General.nominal_frequency / 2.0
	# Defaults to not excluding any data points
	def_impedance_exclude = 0.0

	# Maximum harmonic order (no real impact just avoids excessive loops)
	max_harm = 100

	# Strings defined in inputs
	unlimited_inputs = 'Unlimited'
	custom_inputs = 'Custom'
	min_frequency = 'Minimum Frequency (Hz)'
	max_frequency = 'Maximum Frequency (Hz)'
	percentage_to_exclude = 'Percentage to Exclude (%)'
	max_vertices = 'Maximum No. Vertices'

	# Default values to use for if no loci vertice restrictions are in place
	unlimited_identifier = 10000
	def_max_vertices = unlimited_identifier

	# Minimum allowable number of vertices, less than this and no calculation is really possible
	min_vertices = 4
	max_num_vertices = 100

	# This is the percentage of the impedance that the vertice will be moved by, the smaller this is
	# the longer results processing will take but the less likely that the loci will be increased in
	# size excessively
	vertice_step_size = 0.1


class GuiDefaults:
	gui_title='PSC - Automated PowerFactory Frequency Scans Tool'


	# Default labels for buttons (only those which get changed during running)
	button_select_settings_label = 'Select Settings File'

	# Default extensions used in file type selection windows
	xlsx_types = (('xlsx files', '*.xlsx'), ('All Files', '*.*'))

	font_family = 'Helvetica'
	psc_uk = 'PSC UK'
	psc_phone = '\nPSC UK:  +44 1926 675 851'
	psc_font = ('Calibri', '10', 'bold')
	psc_color_web_blue = '#%02x%02x%02x' % (43, 112, 170)
	psc_color_grey = '#%02x%02x%02x' % (89, 89, 89)
	font_heading_color = '#%02x%02x%02x' % (0, 0, 255)
	img_size_psc = (120, 120)

	img_pth_psc_main = os.path.join(local_directory, 'PSC Logo RGB Vertical.png')
	img_pth_psc_window = os.path.join(local_directory, 'PSC Logo no tag-1200.gif')

	hyperlink_psc_website = 'https://www.pscconsulting.com/'

	# Colors
	color_main_window = '#%02x%02x%02x' % (239, 243, 241)
	# Color of pop-up windows which may be different to main window
	color_pop_up_window = color_main_window
	error_color = '#%02x%02x%02x' % (255, 32, 32)


	# Reference to the Tkinter binding of a mouse button
	mouse_button_1 = '<Button - 1>'

class StudySettings:
	# Names for index in Inputs spreadsheet
	export_folder = 'Results_Export_Folder'
	results_name = 'Excel_Results'
	pf_network_elm = 'Net_Elm'
	pre_case_check = 'Pre_Case_Check'
	delete_created_folders = 'Delete_Created_Folders'
	export_to_excel = 'Export_to_Excel'
	export_rx = 'Excel_Export_RX'
	export_mutual = 'Excel_Export_Z12'
	include_intact = 'Include_Intact'
	include_convex = 'Include_Convex'

	# Base_Scenario columns
	name = 'NAME'
	project = 'Database'
	studycase = 'Studycase'
	scenario = 'Operational Scenario'
	studycase_columns = [name, project, studycase, scenario]

	# Default values
	def_results_name = 'Results_'

class Author:
	""" Contains details of the author """
	developer = 'David Mills'
	email = 'david.mills@PSCconsulting.com'
	phone = '+44 7899 984158'
	company = 'PSC UK'
	website = 'https:\\www.pscconsulting.com'
	contact_summary = '\t{}\n\t{}\n\t{} - {}\n\t{}'.format(
		developer,
		company,
		email, phone,
		website
	)
