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
	:return pd.DataFrame _df:  Return data frame processed ready for exporting to Excel in HAST format
	"""
	c = constants.ResultsExtract

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
	filename = os.path.splitext(os.path.basename(file))[0]
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

def extract_results(pth_file, df, hast_inputs):
	"""
		Extract results into workbook with each result on separate worksheet
	:param str pth_file:  File to save workbook to
	:param pd.DataFrame df:  Pandas dataframe to be extracted
	:param hast2.excel_writing.HASTInputs hast_inputs:  Inputs used for class study
	:return:
	"""
	# Obtain constants
	c = constants.ResultsExtract
	start_row = c.start_row
	# Obtain the variables to export and they will be expoted in this order
	vars_to_export = hast_inputs.vars_to_export()

	# Delete empty column headers which correlate to either frequency of harmonic number data which
	# has already been used as the index
	del df['']

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
	hast_inputs = hast2.excel_writing.HASTInputs(analysis_dict)
	d_terminals = hast_inputs.dict_of_terms
	# study_settings = analysis_dict[constants.HASTInputs.study_settings]

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

	combined_df = pd.concat(dfs, axis=1)

	# Extract results into a spreadsheet
	pth_results = os.path.join(target_dir, 'Results.xlsx')
	extract_results(pth_file=pth_results, df=combined_df, hast_inputs=hast_inputs)

	t2 = time.time()
	logger.warning('Complete process took {:.2f} seconds'.format(t2-t0))




# ----- UNIT TESTS -----
# TODO: Unit tests to be produced