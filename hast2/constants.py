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

class PowerFactory:
	"""
		Constants used in this script
	"""
	sht_Filters = 'Filters'
	sht_Terminals = 'Terminals'
	sht_Scenarios = 'Scenarios'
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
	pf_scenario = 'IntScenarios'
	pf_filter = 'ElmShnt'
	pf_cubicle = 'StaCubic'
	pf_term_voltage = 'uknom'
	pf_shn_voltage = 'ushnm'
	pf_shn_type = 'shtype'
	pf_shn_q = 'qtotn'
	pf_shn_freq = 'fres'
	pf_shn_qfactor = 'greaf0'
	pf_shn_rp = 'rpara'