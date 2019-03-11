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