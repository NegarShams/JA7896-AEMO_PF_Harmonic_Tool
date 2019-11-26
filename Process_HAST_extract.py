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

import os
import pandas as pd
import numpy as np
import glob
import logging
import hast2_1 as hast2
import hast2_1.constants as constants
import time
import collections
import math
import pickle

# Meta Data
__author__ = 'David Mills'
__version__ = '2.1.2'
__email__ = 'david.mills@PSCconsulting.com'
__phone__ = '+44 7899 984158'
__status__ = 'In Development - Beta'
# Get name of current script for use in log messages
__script__ = os.path.basename(__file__)

logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s',
					level=logging.INFO,
					datefmt='%Y-%m-%d %H:%M:%S')
logger = logging.getLogger(constants.logger_name)

# Populate the following to avoid a GUI import
list_of_folders_to_import = []
target_file = None

# Whether to include graphs when exporting and whether to include CONVEX HULL
PLOT_GRAPHS = True
INCLUDE_CONVEX_HULL = False

# For Backwards compatibility
INCLUDE_NOM_VOLTAGE = True

# To resolving naming issues with results extract from PowerFactory

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
	terminal_names = dict_of_terms.values()

	# Process each variable to identify the mutual / terminal names
	for var in vars_list:
		if '.{}'.format(c.pf_mutual) in var:
			# Remove the reference to mutual impedance from the terminal
			var_name = var.replace('.{}'.format(c.pf_mutual), '')

			term1 = None
			term2 = None
			for term in terminal_names:
				if var_name.startswith(term):
					term1 = term
				if var_name.endswith(term):
					term2 = term

			# Combine terminals into name
			ref_terminal = (term1, term2)

			# Check that names have been determined for each variable
			if not all(ref_terminal):
				logger.error(('Not completely sure if the reference terminals for mutual impedance {} are correct.'
								'In determining the reference terminals it is assumed that the mutual impedance is '
								'named by HAST as "Terminal 1_Terminal 2" but this seems not to be the case here.'
								'The following terminals have been used: \n'
								'{}\n{}').format(var_name, ref_terminal[0], ref_terminal[1]))
			# Two mutual names returns in lists for each direction and each ref_terminal
			# Variable names as lists in reverse order
			var_name = ('_'.join(ref_terminal),'_'.join(ref_terminal[::-1]))
			# Mutual name found so exit for loop
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

def process_file_name(file_name, sc_names, cont_names):
	"""
		Splits up the file name to identify the study type, case, contingency and filter details
	:param str file_name:  Existing file name
	:param list sc_names:  List of sc_names considered in study
	:param list cont_names:  List of cont_names considered in study
	:return list components: [study_type, study_case, contingency, filter_name) where file_name remaining
	"""
	c = constants.ResultsExtract
	study_type = ''
	sc_name = ''
	cont_name = ''
	filter_name = ''

	# Find the relevant study type
	for s_type in c.study_types:
		if s_type in file_name:
			study_type = s_type
			file_name = file_name.replace('{}_'.format(study_type), '')
			break

	# Find which study case is shown
	for sc in sc_names:
		if sc in file_name:
			sc_name = sc
			file_name = file_name.replace('{}_'.format(sc_name), '')
			break

	# Find which contingnecy is considered
	for cont in cont_names:
		if cont in file_name:
			cont_name = cont
			file_name = file_name.replace('{}_'.format(cont_name), '')
			break

	file_name_splits = file_name.split('_')
	_, filter_name = split_contingency_filter_values(file_name_splits, starting_point=0)

	return study_type, sc_name, cont_name, filter_name

def manual_adjustments_to_var_names(list_of_var_names, dict_of_adjustments):
	"""

	:param list_of_var_names:
	:param dict_of_adjustments:
	:return:
	"""
	new_var_names = []
	for var in list_of_var_names:
		for mistake, correction in dict_of_adjustments.items():
			if mistake in var:
				var = var.replace(mistake,correction)
				break

		new_var_names.append(var)

	return new_var_names

def process_file(pth_file, hast_inputs, manual_adjustments=dict()):
	"""
		# Process the imported HAST results file into a dataframe with the relevant multi-index
	:param str pth_file:  Full path to results that need importing
	:param hast2_1.excel_writing.HASTInputs hast_inputs:  Handle to the HAST inputs data
	:param dict manual_adjustments:  Contains naming conversions to use for the variables if they need adjusting
	:return pd.DataFrame _df:  Return data frame processed ready for exporting to Excel in HAST format
	"""
	c = constants.ResultsExtract

	# Import dataframe
	_df = pd.read_csv(pth_file, header=[0, 1], index_col=0)

	# Check if index is frequency or harmonic no. and if latter then multiply by nominal frequency
	if not _df.index[0] >= constants.nom_freq:
		_df.index = _df.index * constants.nom_freq
	_df.index.name = c.lbl_Frequency

	# Get from results file:
	#	Node / Mutual Name
	#	Variable type
	filename = os.path.splitext(os.path.basename(pth_file))[0]

	# Process the file name to understand the study case being considered
	study_type, sc_name, cont_name, filter_name = process_file_name(file_name=filename,
																	sc_names=hast_inputs.sc_names,
																	cont_names=hast_inputs.cont_names)
	# Determine if different filter modelling has been included
	if filter_name != '':
		full_name = '{}_{}_{}'.format(sc_name, cont_name, filter_name)
	else:
		full_name = '{}_{}'.format(sc_name, cont_name)

	columns = list(zip(*_df.columns.tolist()))
	# To manually deal wit renaming of mutual impedance values
	if manual_adjustments:
		logger.debug('\t - \t\t Manual adjustment of variable names being completed for file {}'.format(filename))
		var_names = manual_adjustments_to_var_names(columns[0], dict_of_adjustments=manual_adjustments)
	else:
		var_names = columns[0]
	var_types = columns[1]

	df_mutual = pd.DataFrame().reindex_like(_df)
	df_mutual = df_mutual.drop(_df.columns, axis=1)
	new_var_names1 = []
	new_var_names2 = []
	new_ref_terminals1 = []
	new_ref_terminals2 = []
	new_var_types1 = []
	new_var_types2 = []
	for i, var in enumerate(var_names):
		_var_names, _ref_terms = extract_var_name(var_name=var, dict_of_terms=hast_inputs.dict_of_terms)
		_var_type = extract_var_type(var_types[i])
		# Mutual impedance data
		if type(_var_names) is tuple:
			# If mutual impedance data is returned then need to duplicate dataframe to create data in the other direction
			new_var_names1.append(_var_names[0])
			new_var_names2.append(_var_names[1])
			new_ref_terminals1.append(_ref_terms[0])
			new_ref_terminals2.append(_ref_terms[1])
			new_var_types1.append(_var_type)
			new_var_types2.append(_var_type)
			# Create a copy of the data
			df_mutual[_df.columns[i]] = _df[_df.columns[i]]
		else:
			# If no mutual impedance data then just extract variable names in lists
			new_var_names1.append(_var_names)
			new_ref_terminals1.append(_ref_terms)
			new_var_types1.append(_var_type)

	# Produce new multi-index containing new headers
	col_headers1 = [(ref_terminal, var_name, sc_name, cont_name, filter_name, full_name, var_type)
				   for ref_terminal, var_name, var_type in zip(new_ref_terminals1, new_var_names1, new_var_types1)]
	col_headers2 = [(ref_terminal, var_name, sc_name, cont_name, filter_name, full_name, var_type)
					for ref_terminal, var_name, var_type in zip(new_ref_terminals2, new_var_names2, new_var_types2)]

	# Names for headers
	names = (c.lbl_Reference_Terminal,
			 c.lbl_Terminal,
			 c.lbl_StudyCase,
			 c.lbl_Contingency,
			 c.lbl_Filter_ID,
			 c.lbl_FullName,
			 c.lbl_Result)
	# Produce multi-index and assign to DataFrames
	columns1 = pd.MultiIndex.from_tuples(tuples=col_headers1, names=names)
	columns2 = pd.MultiIndex.from_tuples(tuples=col_headers2, names=names)
	# Replace previous multi-index with new
	_df.columns = columns1
	df_mutual.columns = columns2

	# Combine dataframes into one and return
	_df = pd.concat([_df, df_mutual], axis=1)

	# Obtain nominal voltage for each terminal
	idx = pd.IndexSlice
	# Add in row to contain the nominal voltage
	_df.loc[c.idx_nom_voltage] = str()

	dict_nom_voltage = dict()

	# Find the nominal voltage for each terminal (if exists)
	for term, df in _df.groupby(axis=1, level=c.lbl_Reference_Terminal):
		idx_filter = idx[:,:,:,:,:,:,'e:uknom']
		try:
			# Obtain nominal voltage and then set row values appropriately to include in results
			nom_voltage = df.loc[:,idx_filter].iloc[0,0]
			dict_nom_voltage[term] = nom_voltage
			# #_df.loc[c.idx_nom_voltage, idx[term,:,:,:,:,:,:]] = nom_voltage
			# #df.loc[c.idx_nom_voltage, :] = nom_voltage
		except KeyError:
			pass

	# Check for any duplicated multi-index entries (typically contingencies) and rename
	to_keep = 'first'
	if any(_df.columns.duplicated(keep=to_keep)):
		logger.debug('Processing duplicated results in the HAST results file: {}'.format(pth_file))
		# Get duplicated and non-duplicated into separate DataFrames
		duplicated_entries = _df.loc[:, _df.columns.duplicated(keep=to_keep)]
		non_duplicated_entries = _df.loc[:, ~_df.columns.duplicated(keep=to_keep)]

		# Rename duplicated_entries and report to user (contingency is the entry that is renamed)
		# Only allows for a single duplicated entry
		if any(duplicated_entries.columns.duplicated()):
			logger.critical(
				(
					'Unexpected error when trying to deal with duplicate columns for processing of the HAST results '
					'file: {}'
				).format(pth_file)
			)
			raise IOError('Multiple duplicated entries')

		# Produce new column labels for each contingency
		terminals = set(duplicated_entries.columns.get_level_values(level=c.lbl_Reference_Terminal))
		msg = (
			(
				'During processing of the HAST results file: {} some duplicated entries have been for the '
				'following terminals: \n'
			).format(pth_file)
		)
		msg += '\n'.join(['\t - Terminal:  {}'.format(x) for x in terminals])
		msg += '\n To resolve this the following changes have been made to the duplicated entries:\n'
		for label in (c.lbl_Contingency, c.lbl_FullName):
			# Get all the old and new labels and combine together into a lookup dictionary
			old_labels = set(duplicated_entries.columns.get_level_values(level=label))
			new_labels = ['{}({})'.format(x, 1) for x in old_labels]
			replacement = dict(zip(old_labels, new_labels))
			# Replace the duplicated entries with the new ones
			duplicated_entries.rename(columns=replacement, level=label, inplace=True)
			msg += '\n'.join(
				['\t - For {} value {} has been replaced with {}'.format(label, k, v) for k,v in replacement.items()]
			) +'\n'
		logger.warning(msg)

		# Combine the two DataFrames back into a single DataFrame
		_df = pd.concat([duplicated_entries, non_duplicated_entries], axis=1)

	# Update the DataFrame to include the nominal voltage for every terminal
	for term, nom_voltage in dict_nom_voltage.items():
		_df.loc[c.idx_nom_voltage, idx[term,:,:,:,:,:,:]] = nom_voltage

	# Add the new nominal voltage into the index data
	_df = _df.T.set_index(c.idx_nom_voltage, append=True).T
	_df = _df.reorder_levels([c.lbl_Reference_Terminal, c.lbl_Terminal, c.idx_nom_voltage,
							  c.lbl_StudyCase, c.lbl_Contingency, c.lbl_Filter_ID,
							  c.lbl_FullName, c.lbl_Result], axis=1)

	# For backwards compatibility, remove multi-level index if required
	if not INCLUDE_NOM_VOLTAGE:
		cols = _df.columns.droplevel(c.idx_nom_voltage)
		_df.columns = cols

	return _df

def graph_grouping(df, group_by=constants.ResultsExtract.chart_grouping, startcol=0):
	"""
		Determines sizes for grouping of graphs together
		CAVEAT:  Assumes that the DataFrame order matches with the order in Excel
	:param pd.DataFrame df: Dataframe to calculate grouping for
	:param tuple group_by: (optional = constants.ResultsExtract.graph_grouping) = Levels to group by
	:param int startcol: (optional=0) Column which graphs start from to be added to the dataframe columns
	:return dict col_nums:  {Name of graph: Relative column numbers for results
	"""
	# Determine number of results in each group so that it can determine whether multiple graphs are needed
	df_groups = list(df.groupby(axis=1, level=group_by).size())

	# If only single plot on each graph then no need to separate at this level so go up 1 level
	if max(df_groups) == 1 and len(group_by)>1:
		logger.debug('Only single value for each entry so no need to split across multiple graphs')
		group_by = group_by[:-1]

	# Produce dictionary which looks up column numbers for each set of results that are to be grouped by
	col_nums = collections.OrderedDict()
	for key, v in df.groupby(axis=1, level=group_by):
		if type(key) is not str:
			k = '_'.join(key)
		else:
			k = key
		list_col_nums = list(df.columns.get_locs(list(map(list, zip(*v.columns.to_list())))))
		col_nums[k] =[x+startcol for x in list_col_nums]

	return col_nums

def extract_results(pth_file, df, vars_to_export, plot_graphs=True):
	"""
		Extract results into workbook with each result on separate worksheet
	:param str pth_file:  File to save workbook to
	:param pd.DataFrame df:  Pandas dataframe to be extracted
	:param list vars_to_export:  List of variables to export based on Hast Inputs class
	:param bool plot_graphs:  (optional=True) - If set to False then graphs will not be exported
	:return None:
	"""
	logger.info('Exporting imported results to {}'.format(pth_file))
	# Obtain constants
	c = constants.ResultsExtract
	start_row = c.start_row

	# Delete empty column headers which correlate to either frequency of harmonic number data which
	# has already been used as the index
	df.drop(columns=c.lbl_to_delete, inplace=True, level=0)

	# Group the data frame by node name
	list_dfs = df.groupby(level=c.lbl_Reference_Terminal, axis=1)
	num_nodes = len(list_dfs)
	logger.info('Exporting results for {} nodes'.format(num_nodes))

	# Export to excel with a new sheet for each node
	i=0
	try:
		with pd.ExcelWriter(pth_file, engine='xlsxwriter') as writer:
			for node_name, _df in list_dfs:
				logger.info(' - \t {}/{} Exporting node {}'.format(i+1, num_nodes, node_name))
				i += 1
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
							logger.info(' \t - \t Adding graph for node {}'.format(node_name))

							num_rows = df_to_export.shape[0]
							# Get number of columns to include in each graph grouping
							dict_graph_grouping = graph_grouping(df=df_to_export, startcol=col+1)
							names = df_to_export.columns.names
							row_cont = start_row + names.index(constants.ResultsExtract.lbl_FullName)


							add_graph(writer, sheet_name=node_name,
									  row_cont=row_cont,
									  row_start=start_row + len(names) + 1,
									  col_freq=col,
									  num_rows=num_rows,
									  graph_groups=dict_graph_grouping,
									  chrt_row_num=0)

							# Get grouping of graphs to compare study cases
							dict_graph_grouping = graph_grouping(df=df_to_export,
																 group_by=constants.ResultsExtract.chart_grouping_base_case,
																 startcol=col+1)

							add_graph(writer, sheet_name=node_name,
									  row_cont=row_cont,
									  row_start=start_row + len(names) + 1,
									  col_freq=col,
									  num_rows=num_rows,
									  graph_groups=dict_graph_grouping,
									  chrt_row_num=1)

						col = col + df_to_export.shape[1] + c.col_spacing
					else:
						logger.warning('No results imported for variable {} at node {}'.format(var, node_name))
	except PermissionError:
		logger.critical(('Unable to write to excel workbook {} since it is either already open or you do not have '
						 'the appropriate permissions to save here.  Please check and then rerun {}')
						.format(pth_file, __script__))
		raise PermissionError('Unable to write to workbook')

	return None

def split_plots(max_plots, graph_groups):
	"""
		Figures out how to split the plots into groups based on the grouping and maximum of 255 plots (or max_plots)
		Returns the relevant names and excel column number
	:param int max_plots:  Maximum number of plots to include
	:param collections.OrderedDict graph_groups:  Dictionary of graph grouping in the format
		key:[list of relative column numbers]
	List of graph grouping produced by <graph_grouping>
	:return collections.OrderedDict graphs:  Dictionary of the graph title and relevant column numbers
	"""
	graphs = collections.OrderedDict()

	for title, list_of_cols in graph_groups.items():
		# Get all columns in range associated with this group
		# Number of plots this group needs to be split into
		number_of_plots = math.ceil(len(list_of_cols) / max_plots)

		# If greater than 1 then split into equal size groups with basecase plot at starting point
		if number_of_plots > 1:
			steps = math.ceil(len(list_of_cols) / number_of_plots)
			a = [list_of_cols[i * steps:(i + 1) * steps] for i in range(int(math.ceil(len(list_of_cols) / steps)))]

			for i, group in enumerate(a):
				graphs['{}({})'.format(title, i + 1)] = group
		else:
			graphs['{}'.format(title)] = list_of_cols

	return graphs

def add_graph(writer, sheet_name, row_cont, row_start, col_freq, num_rows,
			  graph_groups, chrt_row_num):
	"""
		Add graph to HAST export
	:param pd.ExcelWriter writer:  Handle for the workbook that will be controlling the excel instance
	:param str sheet_name: Name of sheet to add graph to
	:param int row_cont: Row number which contains the contingency description
	:param int row_start: Start row for results to plot
	:param int col_freq: Number of column which contains the frequency data
	:param int num_rows:  Number of rows containing the data to be plotted
	:param collections.OrderedDict graph_groups:  Names and groups to use for graph grouping in the form
			key:[column numbers relative to the first column]
	:param int chrt_row_num:  Number for this chart which determines the vertical row number the chart is added to
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
	charts = []
	max_plots = len(color_map) - 1

	plots = split_plots(max_plots, graph_groups)

	for chart_name, cols in plots.items():
		chrt = wkbk.add_chart(c.chart_type)
		# Adjusted to include the name of the sheet in the chart title and fixing the font_size
		# #chrt.set_title({'name':chart_name})
		chrt.set_title({'name':'{} - {}'.format(sheet_name, chart_name),
						'name_font':{'size':c.font_size_chart_title}})
		charts.append(chrt)
		color_i = 0

		# Loop through each column and add series
		for col in cols:
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

	for i, chrt in enumerate(charts):
		# Add axis labels
		chrt.set_x_axis({'name': c.lbl_Frequency, 'label_position': c.lbl_position,
						 'major_gridlines': c.grid_lines})
		chrt.set_y_axis({'name': c.lbl_Impedance, 'label_position': c.lbl_position,
						 'major_gridlines': c.grid_lines})

		# Set chart size
		chrt.set_size({'width': c.chrt_width, 'height': c.chrt_height})

		# Add the legend to the chart
		sht.insert_chart(c.chrt_row+(c.chrt_vert_space*chrt_row_num), c.chrt_col+(c.chrt_space*i), chrt)

def import_all_results(search_pth, hast_inputs, search_type='FS'):
	"""
		Function to import all results into a single DataFrame
	:param str search_pth: Directory which contains the exported results files which are to be imported
	:param hast2_1.excel_writing.HASTInputs hast_inputs:  Handle to the HAST inputs data
	:param str search_type: (Optional='FS') - Leading characters to use in search string
	:return pd.DataFrame single_df:  Combined imported files into single DataFrame
	"""
	# Get list of all files in folder for frequency scan
	t0 = time.time()
	files = glob.glob('{}\{}*.csv'.format(search_pth, search_type))
	no_files = len(files)
	logger.info('Importing {} hast results files in directory: {}'.format(no_files, search_pth))

	# Import each results file and combine into a single dataframe
	dfs = []
	for i, file in enumerate(files):
		_df = process_file(pth_file=file, hast_inputs=hast_inputs)
		dfs.append(_df)
		logger.info(' - \t {}/{} HAST results file: {} imported'.format(i+1, no_files, os.path.basename(file)))

	if len(dfs) != no_files:
		logger.error(('There was an issue in the file import and not all were imported.\n'
					   'Only {} of {} files were imported\n'
						'However, the script will continue until something critical occurs')
					   .format(len(dfs), no_files))
	t1 = time.time()
	logger.info('{} of {} results files in the folder: {} imported in {:.2f} seconds'
				.format(len(dfs), no_files, search_pth, t1-t0))

	single_df = pd.concat(dfs, axis=1)
	logger.info('Single dataset for all results in folder:  {} produced in {:.2f} seconds'
				.format(search_pth, time.time()-t0))
	return single_df

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
		analysis_dict = excel_cls.import_excel_harmonic_inputs(workbookname=hast_inputs_workbook)

	# Process the imported workbook into
	processed_inputs = hast2.excel_writing.HASTInputs(analysis_dict)
	logger.info('Inputs from HAST file: {} extracted'.format(hast_inputs_workbook))
	return processed_inputs

def combine_multiple_hast_runs(search_pths, drop_duplicates=True):
	"""
		Function will combine multiple HAST results extracts into a single HAST results file
	:param list search_pths:  List of folders which contain the results files to be combined / extracted
	 							each folder must contain raw .csv results exports + a HAST inputs
	:param bool drop_duplicates:  (Optional=True) - If set to False then duplicated columns will be included in the output
	:return pd.DataFrame df, list vars_to_export, bool export_graphs, bool export_convex_hull:
				Combined results into single dataframe,
				list of variables for export
				Whether any of the hast files required graphs to be exported
				Whether any of the hast files required convex hull to be produced
	"""
	t0 = time.time()
	logger.info('Importing all hast results files in list folders: \n' +
				 '\n'.join('\t{}' for  _ in range(len(search_pths))).format(*search_pths))
	# Loop through each folder, obtain the hast files and produce the dataframes
	c = constants.ResultsExtract
	all_dfs = []
	vars_to_export = []

	# Constant used to determine whether graphs should be included or not
	export_graphs = PLOT_GRAPHS
	export_convex_hull = INCLUDE_CONVEX_HULL

	# Loop through each folder, import the hast inputs sheet and results files
	for folder in search_pths:
		t0 = time.time()
		logger.debug('Importing hast files in folder: {}'.format(folder))
		_hast_inputs = get_hast_values(search_pth=folder)
		_combined_df = import_all_results(search_pth=folder,
										  hast_inputs=_hast_inputs)
		all_dfs.append(_combined_df)
		logger.debug('Importing of all results in folder {} completed in {:.2f} seconds'
					 .format(folder, time.time()-t0))

		# Include list of variables for export
		vars_to_export.extend(_hast_inputs.vars_to_export())

		# Determines whether graphs should be plotted or not, if any of the HAST inputs requires graphs
		# or the constant above PLOT_GRAPHS then they will be included
		export_graphs = any((_hast_inputs.export_to_excel, export_graphs))
		export_convex_hull = any((export_convex_hull, _hast_inputs.include_convex_hull))

	t1 = time.time()
	logger.info('All results imported in {:.2f} seconds'.format(t1-t0))

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
		# Remove any duplicate datasets with matching column names and rows
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

		# Check if any columns are still duplicated as this must now be due to different result sets
		# Check and rename results if duplicated study case names at level (Full Results Name)
		duplicated_col_names = df.columns.get_duplicates()
		if not duplicated_col_names.empty:
			# Produce dictionary for any duplicated names
			dict_duplicates = {k: 1 for k in duplicated_col_names}
			idx_sc = df.columns.names.index(c.lbl_StudyCase)
			idx_full_name = df.columns.names.index(c.lbl_FullName)

			# Get names for existing columns
			existing_columns = df.columns.tolist()

			new_cols = []
			for col in existing_columns:
				# Loop through each column and rename those which appear in the list of duplicated columns
				try:
					duplicated_count = dict_duplicates[col]
				except KeyError:
					duplicated_count = False

				if duplicated_count:
					# Rename duplicated columns to be the form 'sc_name(dup_count)'
					new_col = list(col)
					sc_name = col[idx_sc]
					full_name = col[idx_full_name]
					# Produce new names for study case and full name
					new_sc_name = '{}({})'.format(sc_name, duplicated_count)
					new_full_name = full_name.replace(sc_name, new_sc_name)
					new_col[idx_sc] =  new_sc_name
					new_col[idx_full_name] = new_full_name
					dict_duplicates[col] += 1
				else:
					new_col = col

				new_cols.append(tuple(new_col))

			columns=pd.MultiIndex.from_tuples(tuples=new_cols, names=df.columns.names)
			df.columns = columns
			duplicated_col_names2 = df.columns.get_duplicates()
			logger.warning(('Some results have the same study case name but different values, the user should '
							'check the results that are being combined and confirm where the mistake has been made.\n'
							'For now the studycases have been renamed with (1), (2), (etc.) for presentation.\n'
							'In total {} columns have been renamed')
						   .format(len(duplicated_col_names)-len(duplicated_col_names2))
						   )
			if not duplicated_col_names2.empty:
				raise IOError(' There are still duplicated columns being detected')

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

	# Sort the results so that in study_case name order
	df.sort_index(axis=1, level=[c.lbl_Reference_Terminal, c.lbl_Terminal, c.lbl_StudyCase], inplace=True)

	logger.info('Imported results combined and duplicates removed in {:.2f}'.format(time.time()-t1))
	return df, vars_to_export, export_graphs, export_convex_hull

def calc_rx_boundaries(df, sc=None):
	"""
		Script to extract from the Dataframe the necessary R and X data and to calculate the convex hull for each
		harmonic order.
	:param pd.DataFrame df:  Dataframe for combined data from all results from which R and X data will be extracted
	:param list sc:  (sc=None) - Optional name for study case if want to restrict results
	:return pd.DataFrame df_boundaries:  Returns the boundary coordinates for the RX data specific to this ConvexHull
	"""

	# Extract R and X data from dataframe
	idx = pd.IndexSlice
	if sc is not None:
		# If a list of study cases to include has been provided then include that in the search
		idx_search = idx[:,:,:,sc,:,:,:,[constants.PowerFactory.pf_r1, constants.PowerFactory.pf_x1]]
	else:
		idx_search = idx[:,:,:,:,:,:,:,[constants.PowerFactory.pf_r1, constants.PowerFactory.pf_x1]]
	df_rx = df.loc[:,idx_search]

	# Loop through each node and then group into harmonic orders
	# TODO: Not fully implemented, now need to develop code to calculate ConvexHull

def load_extract_dataframe(targ_file, df=None, load=False):
	"""
		Function to load or extract a dataframe to a target path
		as a pickled file.  Used for faster importing when carrying out
		unittesting.
	:param str targ_file:  Full result file name (including extension) where df is saved
	:param pd.DataFrame df: DataFrame to be extracted
	:param bool load:  If set to True then loads rather than saves
	:return pd.DataFrame df:  Returns either newly imported dataframe or previously exported
	"""
	if load:
		with open(targ_file, 'rb') as input_file:
			df = pickle.load(input_file)
	elif df is not None:
		with open(targ_file, 'wb') as output_file:
			pickle.dump(df, output_file)
	else:
		msg = ('Set to save file (load = {} but yet no DataFrame as been provided (df = {}'
			.format(load, df))
		print(msg)
		raise ValueError(msg)
	return df

# Only runs if main script is run
if __name__ == '__main__':
	# Start logger
	time_stamps = [time.time()]
	log_format = '%(asctime)s - %(levelname)s - %(message)s'
	logging.basicConfig(format=log_format, level=logging.INFO)
	logger = logging.getLogger(constants.logger_name)

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
	logger.info('User file selection took {:.2f} seconds'.format(time_stamps[-1]-time_stamps[0]))

	# Group the multiple sets of results together and obtain the inputs necessary for the next stage
	combined_df, vars_in_hast, export_graphs, export_convex_hull = combine_multiple_hast_runs(
		search_pths=list_of_folders_to_import
	)

	time_stamps.append(time.time())
	logger.info('Processing HAST inputs took {:.2f} seconds'.format(time_stamps[-1] - time_stamps[-2]))

	extract_results(pth_file=target_file, df=combined_df, vars_to_export=vars_in_hast,
					plot_graphs=export_graphs)
	time_stamps.append(time.time())
	logger.info('Extracting results took'.format(time_stamps[-1] - time_stamps[-2]))

	logger.info('Results extracted to {}'.format(target_file))

	logger.info('Complete process took {:.2f} seconds'.format(time_stamps[-1] - time_stamps[0]))




