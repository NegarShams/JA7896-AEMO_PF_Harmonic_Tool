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
import collections
import math

logger = logging.getLogger()

# Populate the following to avoid a GUI import
list_of_folders_to_import = []
target_file = None

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

def split_contingency_filter_values(list_of_terms, starting_point=2):
	"""
		Due to the method of export from HAST the contingency and filter details need to be separated
		Assumption is that the naming of the contingency is along the lines of:
			cont_name_fname_freq_mvar
	:param list list_of_terms: List of values from which to obtain the contingency and filter details
	:param int starting_point: Start in list where continency is defined
	:return str cont_name, str filter_name:  Returns a string with the contingency name and filter name
	"""
	c = constants.PowerFactory
	filter_pos = None
	for i, var in enumerate(list_of_terms):
		# Determine if any of the variables contain the frequency
		if c.hz in var:
			filter_pos = i-1
			break
	if filter_pos is None:
		# No filter so merge all terms
		contingency_name = '_'.join(list_of_terms[starting_point:])
		filter_name = ''
	else:
		contingency_name = '_'.join(list_of_terms[starting_point:filter_pos])
		filter_name = '_'.join(list_of_terms[filter_pos:])
	return contingency_name, filter_name

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
	# Separate filter and contingency names
	cont_name, filter_name = split_contingency_filter_values(list_of_terms=file_split)

	if filter_name != '':
		full_name = '{}_{}_{}'.format(sc_name, cont_name, filter_name)
	else:
		full_name = '{}_{}'.format(sc_name, cont_name)

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

def graph_grouping(df, group_by=constants.ResultsExtract.chart_grouping):
	"""
		Determines sizes for grouping of graphs together
	:param pd.DataFrame df: Dataframe to calculate grouping for
	:param tuple group_by: (optional = constants.ResultsExtract.graph_grouping) = Levels to group by
	:return list grouping:  {Name of graph, number of columns for results
	"""
	# Determine number of columns to consider in each graph
	df_grouping = df.groupby(axis=1, level=group_by).size()
	# Obtain index keys and values
	keys = df_grouping.index.tolist()
	values = list(df_grouping)

	# If only single plot on each graph then no need to separate at this level so go up 1 level
	if max(values) == 1:
		logger.debug('Only single value for each entry so no need to split across multiple graphs')
		group_by = group_by[:-1]
		df_grouping = df.groupby(axis=1, level=group_by).size()
		keys = df_grouping.index.tolist()
		values = list(df_grouping)

	if len(keys) == 1:
		keys = [keys]
	# Create a tuple with name of chart followed by value (tuple used to ensure order matches)
	grouping = [('_'.join(k),v) for k,v in zip(keys, values)]
	return grouping

def extract_results(pth_file, df, vars_to_export, plot_graphs=True):
	"""
		Extract results into workbook with each result on separate worksheet
	:param str pth_file:  File to save workbook to
	:param pd.DataFrame df:  Pandas dataframe to be extracted
	:param list vars_to_export:  List of variables to export based on Hast Inputs class
	:param bool plot_graphs:  (optional=True) - If set to False then graphs will not be exported
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
					# Results are sorted in study case then contingency then filter order
					df_to_export.to_excel(writer, merge_cells=True,
										  sheet_name=node_name,
										  startrow=start_row, startcol=col,
										  header=include_index, index_label=False)

					# Add graphs if data is self-impedance
					if var == constants.PowerFactory.pf_z1 and plot_graphs:
						num_rows = df_to_export.shape[0]
						num_cols = df_to_export.shape[1]
						# Get number of columns to include in each graph grouping
						dict_graph_grouping = graph_grouping(df=df_to_export)
						names = df_to_export.columns.names
						row_cont = start_row + names.index(constants.ResultsExtract.lbl_FullName)

						# TODO:  Rather than producing chart for all Z1 data should actually produce for each study case

						add_graph(writer, sheet_name=node_name,
								  num_cols=num_cols,
								  col_start=col+1,
								  row_cont=row_cont,
								  row_start=start_row + len(names) + 1,
								  col_freq=col,
								  num_rows=num_rows,
								  graph_grouping=dict_graph_grouping)

					col = df_to_export.shape[1] + c.col_spacing
				else:
					logger.warning('No results imported for variable {} at node {}'.format(var, node_name))

def split_plots(max_plots, start_col, graph_grouping):
	"""
		TODO:  Add description
	:param max_plots:
	:param start_col:
	:param graph_grouping:
	:return graphs:
	"""
	graphs = collections.OrderedDict()

	for title, num in graph_grouping:
		# Get all columns in range associated with this group
		# Plot 0 removed and added back in as starting plot
		all_cols = list(range(start_col + 1, start_col + num))
		# Number of plots this group needs to be split into
		number_of_plots = math.ceil((num-1) / max_plots)

		# If greater than 1 then split into equal size groups with basecase plot at starting point
		if number_of_plots > 1:
			steps = math.ceil((num-1) / number_of_plots)
			a = [all_cols[i * steps:(i + 1) * steps] for i in range(math.ceil(len(all_cols) / steps))]
			a = [[start_col] + x for x in a]
			for i, group in enumerate(a):
				graphs['{}({})'.format(title, i + 1)] = group
		else:
			graphs['{}'.format(title)] = [start_col] + all_cols

		if all_cols == []:
			start_col = start_col+1
		else:
			start_col = max(all_cols) + 1

	return graphs

def add_graph(writer, sheet_name, num_cols, col_start, row_cont, row_start, col_freq, num_rows,
			  graph_grouping):
	"""
		Add graph to HAST export
	:param pd.ExcelWriter writer:
	:param str sheet_name:
	:param list graph_grouping:  Names and groups to use for graph grouping
	:return:
	"""
	c = constants.ResultsExtract
	color_map = c().get_color_map()

	# Get handles
	wkbk = writer.book
	sht = writer.sheets[sheet_name]

	# Calculate the row number for the end of the dataset
	max_row = row_start+num_rows

	# Loop through each column and add series
	# TODO: Need to split into separate chart if more than 255 series
	color_i = 0
	charts = []
	chrt = None
	chrt_count = 0
	group_i = 0
	group_titles = []
	max_plots = len(color_map) - 1

	plots = split_plots(max_plots, col_start, graph_grouping)

	for chart_name, cols in plots.items():
		chrt = wkbk.add_chart(c.chart_type)
		chrt.set_title({'name':chart_name})
		charts.append(chrt)
		color_i = 0

		# Loop through each column and add series
		for col in cols:


			# # Loop through each column and add new series
			# #for i in range(num_cols):
			# #	col = col_start + i
			# #	# Maximum number of plots exceeded, create a new chart
			# #	if color_i > max_plots or chrt is None:
			# #		chrt = wkbk.add_chart(c.chart_type)
			# #		charts.append(chrt)
			# #		chrt_count += 1
			# #		color_i = 0

			# #else:
				# #chrt = charts[chrt_count]

			chrt.add_series({
				'name': [sheet_name, row_cont, col],
				'categories': [sheet_name, row_start, col_freq, max_row, col_freq],
				'values': [sheet_name, row_start, col, max_row, col],
				'marker': {'type': 'none'},
				'line':  {'color': color_map[color_i],
						  'width': c.line_width}
			})

			# color_i is used to determine the maximum number of plots that can be stored
			color_i += 1

	# TODO:  Need to add chart title to detail the study case being looked at

	for i, chrt in enumerate(charts):
		# Add axis labels
		chrt.set_x_axis({'name': c.lbl_Frequency, 'label_position': c.lbl_position})
		chrt.set_y_axis({'name': c.lbl_Impedance, 'label_position': c.lbl_position})

		chrt.set_size({'width': c.chrt_width, 'height': c.chrt_height})

		# Add the legend to the chart
		sht.insert_chart(c.chrt_row, c.chrt_col+(c.chrt_space*i), chrt)

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
		_combined_df = import_all_results(search_pth=folder,
										 terminals=_hast_inputs.dict_of_terms)
		all_dfs.append(_combined_df)
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
	time_stamps = [time.time()]

	# Load GUI to select files
	# TODO: Add option in GUI whether to include graphs in plots
	if tk and len(list_of_folders_to_import) == 0:
		# Load GUI for user to select files
		gui = hast2.gui.MainGUI(title='Select HAST results for processing')
		list_of_folders_to_import = gui.results_files_list
		target_file = gui.target_file
		if list_of_folders_to_import == [] or target_file == '':
			logger.critical('Missing either folders or a target file for processing from GUI')
			raise IOError('Missing either folders or a target file for processing from GUI')
	elif tk and len(list_of_folders_to_import) > 0 and not target_file is None:
		logger.warning('List of folders to import has been provided and so these have been used instead: \n'
					   '{}'.format(list_of_folders_to_import))
	elif len(list_of_folders_to_import)>0 or target_file is None:
		logger.critical(('Since python module <tkinter> could not be imported the user must manually'
						 'enter the folders which should be searched under the variable {} at the top'
						 'of the script {} located in {}')
						.format(str(list_of_folders_to_import),
								os.path.basename(__file__),
								os.path.dirname(__file__)))
		raise IOError('No folders provided for data import')

	time_stamps.append(time.time())
	logger.debug('User file selection took {:.2f} seconds'.format(time_stamps[-1]-time_stamps[0]))

	combined_df, vars_in_hast = combine_multiple_hast_runs(search_pths=list_of_folders_to_import)

	time_stamps.append(time.time())
	logger.debug('Processing HAST inputs took {:.2f} seconds'.format(time_stamps[-1] - time_stamps[-2]))

	extract_results(pth_file=target_file, df=combined_df, vars_to_export=vars_in_hast)
	time_stamps.append(time.time())
	logger.debug('Extracting results took'.format(time_stamps[-1] - time_stamps[-2]))

	logger.info('Results extracted to {}'.format(target_file))

	logger.info('Complete process took {:.2f} seconds'.format(time_stamps[-1] - time_stamps[0]))




