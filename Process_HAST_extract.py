"""
#######################################################################################################################
###											HAST_V2_1																###
###		Script deals with processing of results that have previously been exported into the same format that is 	###
###		used by the rest of the HAST tools for post processing of results											###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###		project JI6973 for EirGrid project PSPF010 - Specialise Support in Power Quality Analysis during 2018		###
###																													###
#######################################################################################################################

DEVELOPMENT CODE:  This code is still in development since has not been produced to account for use of Harmonic Load
Flow or extraction of ConvexHull data to excel.

"""

# IMPORT SOME PYTHON MODULES
# --------------------------------------------------------------------------------------------------------------------
import os
import sys
import pandas as pd
import unittest
import glob
import logging
import hast2_1 as hast2
import hast2_1.constants as constants
import time

target_dir = r'C:\Users\david\Desktop\19_03_28_12_22_12'

def extract_var_name(var_name, dict_of_terms):
	"""
		Function extracts the variable name from the list
	:param str var_name: Name to extract relevant component from
	:param dict dict_of_terms:  Dictionary for looking up HAST name from substation / terminal should be in the form:
		{(Substation.ElmSubstat, Terminal.ElmTerm) : HAST substation}
	:return str var_name: Shortened name to use for processing (var1 = Substation or Mutual, var2 = Terminal)
	"""
	# Variable declarations
	c = constants.PowerFactory
	var1 = False
	var2 = False
	ref_terminal = ''

	# Separate PowerFactory path into individual entries
	vars = var_name.split('\\')

	# Process each variable to identify the mutual / terminal names
	for var in vars:
		if '.{}'.format(c.pf_mutual) in var:
			# Mutual name found so exit for loop
			var_name = var.strip('.{}'.format(c.pf_mutual))
			ref_terminal = var_name.split('_')[0]
			break
		elif '.{}'.format(c.pf_substation) in var:
			var1 = var
		elif '.{}'.format(c.pf_terminal) in var:
			var2 = var

	# Lookup HAST terminal name from input spreadsheet
	try:
		var_name = dict_of_terms[(var1, var2)]
		ref_terminal = var_name
	except KeyError:
		pass

	return var_name, ref_terminal

def extract_var_type(var_type):
	"""
		Function extracts the variable type by splitting at the first space
	:param str var_type: Typically provided in the format 'c:R_12 in Ohms'
	:return str (var_type): Shortened name to use for processing
	"""
	return var_type.split(' ')[0]

def process_file(pth_file, dict_of_terms):
	"""
		# Process the imported HAST results file into a dataframe with the relevant multi-index
	:param str pth_file:  Full path to results that need importing
	:param dict dict_of_terms:  Dictionary of terminals that need updating
	:return:
	"""
	# Import dataframe
	df = pd.read_csv(pth_file, header=[0, 1], index_col=0)

	# Get from file name
	#	Study Case
	#	Contingency
	# 	Filter details?

	# Get from results file:
	#	Node / Mutual Name
	#	Variable type
	filename = os.path.splitext(os.path.basename(file))[0]
	file_split = filename.split('_')
	study_type = file_split[0]
	sc_name = file_split[1]
	cont_name = '_'.join(file_split[2:])
	full_name = '{}_{}'.format(sc_name, cont_name)
	# TODO: Processing filter name from results
	filter_name = ''



	columns = list(zip(*df.columns.tolist()))
	var_names = columns[0]
	var_types = columns[1]

	var_names = [extract_var_name(var, dict_of_terms) for var in var_names]
	var_names, ref_terminals = zip(*var_names)
	var_types = [extract_var_type(var) for var in var_types]
	# Combine into a list
	# var_name_type = list(zip(var_names, var_types))

	# Produce new multi-index containing new headers
	c = constants.ResultsExtract
	col_headers = [(ref_terminal, var_name, sc_name, cont_name, filter_name, full_name, var_type)
				   for ref_terminal, var_name, var_type in zip(ref_terminals, var_names, var_types)]
	names = (c.lbl_Reference_Terminal,
			 c.lbl_Terminal,
			 c.lbl_StudyCase,
			 c.lbl_Contingency,
			 c.lbl_Filter_ID,
			 c.lbl_FullName,
			 c.lbl_Result)
	columns = pd.MultiIndex.from_tuples(tuples=col_headers, names=names)
	# Replace previous multi-index with new
	df.columns = columns

	return df

def process_terminals(list_of_terms):
	"""
		Processes the terminals from the HAST input into a dictionary so can lookup the name to use based on
		substation and terminal
	:param list list_of_terms: List of terminals from HAST inputs, expected in the form
	:return dict dict_of_terms: Dictionary of terminals in the form {(sub name, term name) : HAST name}
	"""
	dict_of_terms = {(k[1], k[2]) : k[0] for k in list_of_terms}
	return dict_of_terms

def extract_results(pth_file, df):
	"""
		Extract results into workbook with each result on separate worksheet
	:param pth_file:
	:return:
	"""
	list_dfs = df.groupby(level=0, axis=1)
	with pd.ExcelWriter(pth_file) as writer:
		for _df in list_dfs:
			if _df[0] != '':
				_df[1].to_excel(writer, sheet_name=_df[0])


# Only runs if main script is run
if __name__ == '__main__':
	# Start logger
	t0 = time.time()
	logger = logging.getLogger()

	# Get and import relevant HAST file to obtain relevant terminal names
	HAST_workbook = glob.glob('{}\HAST Inputs*.xlsx'.format(target_dir))[0]
	with hast2.excel_writing.Excel(print_info=logger.info, print_error=logger.error) as excel_cls:
		# TODO:  Performance improvement possible by speeding up processing of HAST inputs workbook
		analysis_dict = excel_cls.import_excel_harmonic_inputs(workbookname=HAST_workbook)
	# Process the terminals into a lookup dict based
	d_terminals = process_terminals(analysis_dict["Terminals"])  # Uses the list of Terminals
	logger.debug('Processing HAST inputs took {:.2f} seconds'.format(time.time() - t0))

	# Get list of all files in folder for frequency scan
	files = glob.glob('{}\FS*.csv'.format(target_dir))

	# Import each results file and combine into a single dataframe
	dfs = []
	for file in files:
		t1 = time.time()
		df = process_file(file, d_terminals)
		dfs.append(df)

		logger.warning('Processing file {} took {:.2f} seconds'.format(file, time.time() - t1))

	df = pd.concat(dfs, axis=1)

	# Extract results into a spreadsheet
	pth_results = os.path.join(target_dir, 'Results.xlsx')
	# TODO:  To be completed so that looks the same as a HAST export
	extract_results(pth_file=pth_results, df=df)

	t2 = time.time()
	logger.warning('Complete process took {:.2f} seconds'.format(t2-t0))




# ----- UNIT TESTS -----
# TODO: Unit tests to be produced