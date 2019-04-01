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

# Meta Data
__author__ = 'David Mills'
__email__ = 'david.mills@pscconsulting.com'
__phone__ = '+44 7899 984158'
__status__ = 'Constants'

nom_freq = 50.0

# When parallel processing will ensure this number of cpus are kept free
cpu_keep_free = 1
# If unable to calculate this is the assumed maximum number of processes
default_max_processes = 3

class PowerFactory:
	"""
		Constants used in this script
	"""
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
	Filter_type = {'C-Type':4,
				   'Single':0}
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
	# Labels used for frequency scan results extract
	lbl_StudyCase = 'Study Case'
	lbl_Frequency = 'Frequency in Hz'
	lbl_Filter_ID = 'Filter Details'
	lbl_Contingency = 'Contingency'
	lbl_FullName = 'Full Result Name'
	lbl_Terminal = 'Terminal Name'
	lbl_Reference_Terminal = 'Terminal'
	lbl_Result = 'Result Type'
	# Location of m:R, m:X, m:Z, etc.
	loc_pf_variable = 4
	# Location of m:R12, m:X12, m:Z12, etc.
	loc_pf_variable_mutual = loc_pf_variable + 1
	loc_contingency = 1

class HASTInputs:
	base_case = 'Base_Case'
	mutual_variables = ["c:Z_12", "c:R_12", "c:X_12"]
	fs_term_variables = ["m:R", "m:X", "m:Z", "m:phiz"]
	hldf_term_variables = ['m:HD', 'm:THD']

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