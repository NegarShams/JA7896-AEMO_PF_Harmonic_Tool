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
import random

logger = logging.getLogger()
list_of_folders_to_import = []

try:
	import tkinter as tk
except ImportError:
	tk = None
	logger.warning('Unable to import tkinter for use of GUI, you will need to enter files manually')

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
	var_sub = False
	var_term = False
	ref_terminal = ''

	# Separate PowerFactory path into individual entries
	vars_list = var_name.split('\\')

	# Process each variable to identify the mutual / terminal names
	for var in vars_list:
		if '.{}'.format(c.pf_mutual) in var:
			# Mutual name found so exit for loop
			var_name = var.strip('.{}'.format(c.pf_mutual))
			ref_terminal = var_name.split('_')[0]
			break
		elif '.{}'.format(c.pf_substation) in var:
			var_sub = var
		elif '.{}'.format(c.pf_terminal) in var:
			var_term = var
		elif '.{}'.format(c.pf_results) in var:
			# This correlates to the frequency data but that has already been provided and so can
			# be deleted from the results
			var_name = constants.ResultsExtract.lbl_to_delete
			ref_terminal = constants.ResultsExtract.lbl_to_delete
			break

	# Lookup HAST terminal name from input spreadsheet
	if ref_terminal == '':
		try:
			var_name = dict_of_terms[(var_sub, var_term)]
			ref_terminal = var_name
		except KeyError:
			raise KeyError(('The substation and terminal combination [{}, {}] does not appear in the '
							'dictionary of terminals from the HAST spreadsheet')
						   .format(var_sub, var_term))
			pass

	return var_name, ref_terminal

def extract_var_type(var_type):
	"""
		Function extracts the variable type by splitting at the first space
	:param str var_type: Typically provided in the format 'c:R_12 in Ohms'
	:return str (var_type): Shortened name to use for processing
	"""
	var_extract = var_type.split(' ')[0]

	# Raises an exception if poor data inputs given
	if var_extract not in constants.HASTInputs.all_variable_types:
		raise IOError('The variable extracted {} from {} is not one of the input types {}'
					  .format(var_extract, var_type, constants.HASTInputs.all_variable_types))
	return var_extract

def process_file(pth_file, dict_of_terms):
	"""
		# Process the imported HAST results file into a dataframe with the relevant multi-index
	:param str pth_file:  Full path to results that need importing
	:param dict dict_of_terms:  Dictionary of terminals that need updating
	:return pd.DataFrame _df:  Return data frame processed ready for exporting to Excel in HAST format
	"""
	c = constants.ResultsExtract
	# TODO:  Need to account for mutual impedance data to provide in both directions from the single direction produced by HAST

	# Import dataframe
	_df = pd.read_csv(pth_file, header=[0, 1], index_col=0)

	# Check if index is frequency or harmonic no. and if latter then multiply by nominal frequency
	if not _df.index[0] >= constants.nom_freq:
		_df.index = _df.index * constants.nom_freq
	_df.index.name = c.lbl_Frequency

	# Get from file name
	#	Study Case
	#	Contingency
	# 	Filter details?

	# Get from results file:
	#	Node / Mutual Name
	#	Variable type
	filename = os.path.splitext(os.path.basename(pth_file))[0]
	file_split = filename.split('_')
	study_type = file_split[0]
	sc_name = file_split[1]
	cont_name = '_'.join(file_split[2:])
	full_name = '{}_{}'.format(sc_name, cont_name)
	# TODO: Processing filter name from results
	filter_name = ''

	columns = list(zip(*_df.columns.tolist()))
	var_names = columns[0]
	var_types = columns[1]

	var_names = [extract_var_name(var, dict_of_terms) for var in var_names]
	var_names, ref_terminals = zip(*var_names)
	var_types = [extract_var_type(var) for var in var_types]
	# Combine into a list
	# var_name_type = list(zip(var_names, var_types))

	# Produce new multi-index containing new headers
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
	_df.columns = columns

	return _df

def extract_results(pth_file, df, vars_to_export):
	"""
		Extract results into workbook with each result on separate worksheet
	:param str pth_file:  File to save workbook to
	:param pd.DataFrame df:  Pandas dataframe to be extracted
	:param list vars_to_export:  List of variables to export based on Hast Inputs class
	:return:
	"""
	# Obtain constants
	c = constants.ResultsExtract
	start_row = c.start_row

	# Delete empty column headers which correlate to either frequency of harmonic number data which
	# has already been used as the index
	del df[c.lbl_to_delete]

	# Group the data frame by node name
	list_dfs = df.groupby(level=c.lbl_Reference_Terminal, axis=1)

	# Export to excel with a new sheet for each node
	with pd.ExcelWriter(pth_file, engine='xlsxwriter') as writer:
		for node_name, _df in list_dfs:
			col = c.start_col
			for var in vars_to_export:
				# Will only include index and header labels if True
				# include_index = col <= c.start_col
				include_index = True


				df_to_export = _df.loc[:, _df.columns.get_level_values(level=c.lbl_Result)==var]
				if not df_to_export.empty:
					# TODO:  Need to sort in study_case order and then some sort of order for contingencies
					# TODO:  Ideal sort order would be to use same order as appear in hast_inputs
					df_to_export.to_excel(writer, merge_cells=True,
										  sheet_name=node_name,
										  startrow=start_row, startcol=col,
										  header=include_index, index_label=False)

					# Add graphs if data is self-impedance
					if var == constants.PowerFactory.pf_z1:
						num_rows = df_to_export.shape[0]
						num_cols = df_to_export.shape[1]
						names = df_to_export.columns.names
						row_cont = start_row + names.index(constants.ResultsExtract.lbl_FullName)

						# TODO:  Rather than producing chart for all Z1 data should actually produce for each study case

						add_graph(writer, sheet_name=node_name,
								  num_cols=num_cols,
								  col_start=col+1,
								  row_cont=row_cont,
								  row_start=start_row + len(names) + 1,
								  col_freq=col,
								  num_rows=num_rows)

					col = df_to_export.shape[1] + c.col_spacing
				else:
					logger.warning('No results imported for variable {} at node {}'.format(var, node_name))

def add_graph(writer, sheet_name, num_cols, col_start, row_cont, row_start, col_freq, num_rows):
	"""
		Add graph to HAST export
	:param pd.ExcelWriter writer:
	:param str sheet_name:
	:return:
	"""
	c = constants.ResultsExtract
	color_map = c().get_color_map()

	# Get handles
	wkbk = writer.book
	sht = writer.sheets[sheet_name]
	chart = wkbk.add_chart({'type': 'scatter'})

	# Calculate the row number for the end of the dataset
	max_row = row_start+num_rows

	# Loop through each column and add series
	# TODO: Need to split into separate chart if more than 255 series
	color_i = 0
	for i in range(num_cols):
		col = col_start + i
		chart.add_series({
			'name': [sheet_name, row_cont, col],
			'categories': [sheet_name, row_start, col_freq, max_row, col_freq],
			'values': [sheet_name, row_start, col, max_row, col],
			'marker': {'type': 'none'},
			'line':  {'color': color_map[color_i]}
		})

		# color_i is used to determine the color of the plot to use, resetting to 0 if exceeds length
		if color_i >= len(color_map):
			color_i = 0
		else:
			color_i += 1

	# TODO:  Need to add chart title to detail the study case being looked at

	# Add axis labels
	chart.set_x_axis({'Name': c.lbl_Frequency})
	chart.set_y_axis({'Name': c.lbl_Impedance})

	# Add the legend to the chart
	# TODO:  Need to determine position or potentially add position wh
	sht.insert_chart('A1', chart)

def import_all_results(search_pth, terminals, search_type='FS'):
	"""
		Function to import all results into a single DataFrame
	:param str search_pth: Directory which contains the exported results files which are to be imported
	:param dict terminals: Dictionary of terminals and associated values for lookup
	:param str search_type: (Optional='FS') - Leading characters to use in search string
	:return pd.DataFrame combined_df:  Combined imported files into single DataFrame
	"""
	# Get list of all files in folder for frequency scan
	files = glob.glob('{}\{}*.csv'.format(search_pth, search_type))

	# Import each results file and combine into a single dataframe
	dfs = []
	for file in files:
		_df = process_file(pth_file=file, dict_of_terms=terminals)
		dfs.append(_df)

	combined_df = pd.concat(dfs, axis=1)
	return combined_df

def get_hast_values(search_pth):
	"""
		Function to import the HAST inputs and return the class reference needed
	:param str search_pth:  Directory which contains the HAST Inputs
	:return hast2.excel_writing.HASTInputs processed_inputs:  Processed import
	"""
	# Obtain reference to HAST workbook from target directory
	c = constants.HASTInputs

	list_of_input_files = glob.glob('{}\{}*{}'.format(search_pth, c.file_name, c.file_format))
	if len(list_of_input_files) == 0:
		logger.critical(('No HAST inputs file formatted as {}*{} found in the folder {}, please check'
						'a HAST inputs file exists')
						.format(c.file_name, c.file_format, search_pth))
		raise IOError('No HAST Inputs file found')
	elif len(list_of_input_files) > 1:
		hast_inputs_workbook = list_of_input_files[0]
		logger.warning(('Multiple HAST input files were found in the folder {} with the format {}*{} \n'
						'as follows: {} \n'
						'The following file was used assumed to be the correct one: {}')
					   .format(search_pth, c.file_name, c.file_format,
							   list_of_input_files, hast_inputs_workbook))
	else:
		hast_inputs_workbook = list_of_input_files[0]

	# Import HAST workbook using excel_writing import
	with hast2.excel_writing.Excel(print_info=logger.info, print_error=logger.error) as excel_cls:
		# TODO:  Performance improvement possible by speeding up processing of HAST inputs workbook
		analysis_dict = excel_cls.import_excel_harmonic_inputs(workbookname=hast_inputs_workbook)

	# Process the imported workbook into
	processed_inputs = hast2.excel_writing.HASTInputs(analysis_dict)
	return processed_inputs

def combine_multiple_hast_runs(search_pths, drop_duplicates=True):
	"""
		Function will combine multiple HAST results extracts into a single HAST results file
	:param list search_pths:  List of folders which contain the results files to be combined / extracted
	 							each folder must contain raw .csv results exports + a HAST inputs
	:param bool drop_duplicates:  (Optional=True) - If set to False then duplicated columns will be included in the output
	:return pd.DataFrame df, list vars_to_export:  Combined results into single dataframe, list of variables for export
	"""
	# Loop through each folder, obtain the hast files and produce the dataframes
	all_dfs = []
	vars_to_export = []

	# Loop through each folder, import the hast inputs sheet and results files
	for folder in search_pths:
		t0 = time.time()
		logger.debug('Importing hast files in folder: {}'.format(folder))
		_hast_inputs = get_hast_values(search_pth=folder)
		combined_df = import_all_results(search_pth=folder,
										 terminals=_hast_inputs.dict_of_terms)
		all_dfs.append(combined_df)
		logger.debug('Importing of all results in folder {} completed in {:.2f} seconds'
					 .format(folder, time.time()-t0))

		# Include list of variables for export
		vars_to_export.extend(_hast_inputs.vars_to_export())

	# Combine all results together
	df = pd.concat(all_dfs, axis=1)
	# Sorts to improve performance
	df.sort_index(axis=1, level=0, inplace=True)
	# df.sort_index(axis=0, level=0)

	# Create unique list of variables to export without upsetting order
	# 	https://stackoverflow.com/questions/480214/how-do-you-remove-duplicates-from-a-list-whilst-preserving-order
	seen = set()
	seen_add = seen.add
	vars_to_export = [x for x in vars_to_export if not (x in seen or seen_add(x))]

	if drop_duplicates:
		# Remove any duplicate datasets based on column names
		original_shape = df.shape
		# Gets list of duplicated index values and will only compare actual values if duplicated
		df_t = df.T
		cols = df_t.index[df_t.index.duplicated()].tolist()

		if len(cols) != 0:
			# Adds index to columns so only rows which are completely duplicated are considered
			df_t['TEMP'] = df_t.index
			# Removes any completely duplicated rows in transposed dataframe
			# These will be columns which contain exactly the same values in both sets of results
			df_t = df_t.drop_duplicates()
			# Remove temporary index and transpose back
			del df_t['TEMP']
			df = df_t.T

		# Check for changes and record differences
		new_shape = df.shape
		if new_shape[1] != original_shape[1]:
			logger.warning(('The input datasets had duplicated columns and '
						   'therefore some have been removed.\n'
						   '{} columns have been removed')
						   .format(original_shape[1]-new_shape[1]))

		else:
			logger.debug('No duplicated data in results files imported from: {}'
						 .format(search_pths))
		if new_shape[0] != original_shape[0]:
			raise SyntaxError('There has been an error in the processing and some rows have been deleted.'
							  'Check the script')
	else:
		logger.debug('No check for duplicates carried out')

	return df, vars_to_export


# Only runs if main script is run
if __name__ == '__main__':
	# Start logger
	t0 = time.time()

	if tk:
		# Load GUI for user to select files
		pass
	elif len(list_of_folders_to_import)>0:
		logger.critical(('Since python module <tkinter> could not be imported the user must manually'
						 'enter the folders which should be searched under the variable {} at the top'
						 'of the script {} located in {}')
						.format(str(list_of_folders_to_import),
								os.path.basename(__file__),
								os.path.dirname(__file__)))
		raise IOError('No folders provided for data import')


	# Get and import relevant HAST file to obtain relevant terminal names
	# HAST_workbook = glob.glob('{}\HAST Inputs*.xlsx'.format(target_dir))[0]
	# with hast2.excel_writing.Excel(print_info=logger.info, print_error=logger.error) as excel_cls:
	# 	# TODO:  Performance improvement possible by speeding up processing of HAST inputs workbook
	# 	analysis_dict = excel_cls.import_excel_harmonic_inputs(workbookname=HAST_workbook)

	# Process the terminals into a lookup dict based
	# hast_inputs = hast2.excel_writing.HASTInputs(analysis_dict)
	hast_inputs = get_hast_values(search_pth=target_dir)
	d_terminals = hast_inputs.dict_of_terms
	# study_settings = analysis_dict[constants.HASTInputs.study_settings]

	logger.debug('Processing HAST inputs took {:.2f} seconds'.format(time.time() - t0))

	# Get list of all files in folder for frequency scan
	# files = glob.glob('{}\FS*.csv'.format(target_dir))

	# Import each results file and combine into a single dataframe
	# dfs = []
	# for file in files:
	# 	t1 = time.time()
	# 	df = process_file(file, d_terminals)
	# 	dfs.append(df)
	#
	# 	logger.warning('Processing file {} took {:.2f} seconds'.format(file, time.time() - t1))

	# combined_df = pd.concat(dfs, axis=1)
	df = import_all_results(search_pth=target_dir, terminals=d_terminals)

	# Extract results into a spreadsheet
	pth_results = os.path.join(target_dir, 'Results.xlsx')
	# Obtain the variables to export and they will be exported in this order
	vars_in_hast = hast_inputs.vars_to_export()
	extract_results(pth_file=pth_results, df=df, vars_to_export=vars_in_hast)

	t2 = time.time()
	logger.warning('Complete process took {:.2f} seconds'.format(t2-t0))




