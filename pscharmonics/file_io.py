"""
#######################################################################################################################
###											Excel Writing															###
###		Script deals with writing of data to excel and ensuring that a new instance of excel is used for processing	###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###																													###
#######################################################################################################################
"""

import os
import numpy as np
import itertools
import pscharmonics.constants as constants
import glob
import pandas as pd
import collections
import math
import shutil
import time

def update_duplicates(key, df):
	"""
		Function will look for any duplicates in a particular column and then append a number to everything after
		the first one
	:param str key:  Column to lookup
	:param pd.DataFrame df: DataFrame to be processed
	:return pd.DataFrame, bool df_updated, updated: Updated DataFrame and status flag to show updated for log messages
	"""
	# Empty list and initialised value to show no changes
	dfs = list()
	updated = False

	# Group data frame by key value
	# TODO: Do we need to ensure the order of the overall DataFrame remains unchanged
	for _, df_group in df.groupby(key):
		# Extract all key values into list and then loop through to append new value
		names = list(df_group.loc[:, key])
		if len(names) > 1:
			updated = True
			# Check if number of entries is greater than 1, i.e. duplicated
			for i in range(1, len(names)):
				names[i] = '{}({})'.format(names[i], i)

		# Produce new DataFrame with updated names and add to list
		# (copy statement to avoid updating original dataframe during loop)
		df_updated = df_group.copy()
		df_updated.loc[:, key] = names
		dfs.append(df_updated)

	# Combine returned DataFrames
	df_updated = pd.concat(dfs)
	# Ensure order remains as originally
	df_updated.sort_index(axis=0, inplace=True)

	return df_updated, updated

def delete_old_files(pth, logger, thresholds=constants.General.file_number_thresholds):
	"""
		Counts the number of files in a folder and if greater than a certain number it will warn the
		user, if greater than another number it will delete the oldest files.
	:param str pth:  Folder that contains the files
	:param logging.getLogger, logger:  Handle to the logger
	:param tuple thresholds: (int, int) (warning, delete) where
								warning is the number at which a warning message is displayed
								delete is where deleting of the files starts
	:return int num_deleted:
	"""
	thres_warning = thresholds[0]
	thres_delete = thresholds[1]

	num_deleted = 0
	# Check path exists first
	if not os.path.isdir(pth):
		logger.warning(
			(
				'Attempted to check for number of files in folder {} but that folder does not exist and could create '
				'issues later in script'
			).format(pth)
		)

	else:
		# Get number of files
		num_files = len([name for name in os.listdir(pth) if os.path.isfile('{}/{}'.format(pth, name))])

		# Check if more than warning level
		if thres_warning < num_files < thres_delete:
			num_deleted = 0
			logger.warning(
				(
					'There are {} files in the folder {}, when this number gets to {} the oldest will be deleted'
				).format(num_files, pth, thres_delete)
			)
		# Check if more than delete level
		elif num_files >= thres_delete:
			logger.warning(
				(
					'There are {} files in the folder {}, the oldest will be reduced to get this number down to {}'
				).format(num_files, pth, thres_warning)
			)
			num_to_delete = num_files - thres_warning

			list_of_files = os.listdir(pth)
			full_path = ['{}/{}'.format(pth, x) for x in list_of_files]
			for x in range(num_to_delete):
				oldest_file = min(full_path, key=os.path.getctime)
				os.remove(oldest_file)
				full_path.remove(oldest_file)
				num_deleted += 1

		else:
			logger.debug(
				(
					'{} files in folder {} is less than the warning threshold {} and delete thresdhold {}'
				).format(num_files, pth, thres_warning, thres_delete)
			)

	return num_deleted

class ExtractResults:
	def __init__(self, target_file, search_pths):
		""" Process the extraction of the results """
		self.logger = constants.logger

		df, vars = self.combine_multiple_runs(search_pths=search_pths)
		self.extract_results(pth_file=target_file, df=df, vars_to_export=vars)


	def combine_multiple_runs(self, search_pths, drop_duplicates=True):
		"""
			Function will combine multiple results extracts into a single results file
		:param list search_pths:  List of folders which contain the results files to be combined / extracted
									each folder must contain raw .csv results exports + a inputs
		:param bool drop_duplicates:  (Optional=True) - If set to False then duplicated columns will be included in the output
		:return pd.DataFrame df, list vars_to_export:
					Combined results into single dataframe,
					list of variables for export
		"""
		c = constants.Results
		logger = constants.logger

		logger.info(
			'Importing all results files in following list of folders: \n\t{}'.format('\n\t'.join(search_pths))
		)

		# Loop through each folder, obtain the files and produce the dataframes

		all_dfs = []
		vars_to_export = []

		# Loop through each folder, import the inputs sheet and results files
		for folder in search_pths:
			# Import results into a single dataframe
			combined = PreviousResultsExport(pth=folder)
			all_dfs.append(combined.df)

			# Include list of variables for export
			vars_to_export.extend(combined.inputs.settings.get_vars_to_export())

		# Combine all results together
		df = pd.concat(all_dfs, axis=1)
		# Sorts to improve performance
		df.sort_index(axis=1, level=0, inplace=True)

		# Create unique list of variables to export without upsetting order
		# 	https://stackoverflow.com/questions/480214/how-do-you-remove-duplicates-from-a-list-whilst-preserving-order
		seen = set()
		seen_add = seen.add
		vars_to_export = [x for x in vars_to_export if not (x in seen or seen_add(x))]

		if drop_duplicates:
			# Remove any duplicate data sets with matching column names and rows
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

		return df, vars_to_export

	def extract_results(self, pth_file, df, vars_to_export, plot_graphs=True):
		"""
			Extract results into workbook with each result on separate worksheet
		:param str pth_file:  File to save workbook to
		:param pd.DataFrame df:  Pandas dataframe to be extracted
		:param list vars_to_export:  List of variables to export based on Inputs class
		:param bool plot_graphs:  (optional=True) - If set to False then graphs will not be exported
		:return None:
		"""

		self.logger.info('Exporting imported results to {}'.format(pth_file))

		# Obtain constants
		c = constants.Results

		start_row = c.start_row

		# Delete empty column headers which correlate to either frequency of harmonic number data which
		# has already been used as the index
		df.drop(columns=c.lbl_to_delete, inplace=True, level=0)

		# Group the data frame by node name
		list_dfs = df.groupby(level=c.lbl_Reference_Terminal, axis=1)
		num_nodes = len(list_dfs)
		self.logger.info('Exporting results for {} nodes'.format(num_nodes))

		# Export to excel with a new sheet for each node
		i=0
		try:
			with pd.ExcelWriter(pth_file, engine='xlsxwriter') as writer:
				for node_name, _df in list_dfs:
					self.logger.info(' - \t {}/{} Exporting node {}'.format(i+1, num_nodes, node_name))
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
								self.logger.info(' \t - \t Adding graph for node {}'.format(node_name))

								num_rows = df_to_export.shape[0]
								# Get number of columns to include in each graph grouping
								dict_graph_grouping = self.graph_grouping(df=df_to_export, startcol=col+1)
								names = df_to_export.columns.names
								row_cont = start_row + names.index(constants.Results.lbl_FullName)


								self.add_graph(writer, sheet_name=node_name,
										  row_cont=row_cont,
										  row_start=start_row + len(names) + 1,
										  col_freq=col,
										  num_rows=num_rows,
										  graph_groups=dict_graph_grouping,
										  chrt_row_num=0)

								# Get grouping of graphs to compare study cases
								dict_graph_grouping = self.graph_grouping(
									df=df_to_export, group_by=constants.Results.chart_grouping_base_case,
									startcol=col+1
								)

								self.add_graph(writer, sheet_name=node_name,
										  row_cont=row_cont,
										  row_start=start_row + len(names) + 1,
										  col_freq=col,
										  num_rows=num_rows,
										  graph_groups=dict_graph_grouping,
										  chrt_row_num=1)

							col = col + df_to_export.shape[1] + c.col_spacing
						else:
							self.logger.warning('No results imported for variable {} at node {}'.format(var, node_name))
		except PermissionError:
			self.logger.critical(
				(
					'Unable to write to excel workbook {} since it is either already open or you do not have the '
					'appropriate permissions to save here.  Please check and then rerun.'
				).format(pth_file)
			)
			raise PermissionError('Unable to write to workbook')

		return None

	def add_graph(self, writer, sheet_name, row_cont, row_start, col_freq, num_rows,
				  graph_groups, chrt_row_num):
		"""
			Add graph to export
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
		c = constants.Results
		color_map = c().get_color_map()

		# Get handles
		wkbk = writer.book
		sht = writer.sheets[sheet_name]

		# Calculate the row number for the end of the dataset
		max_row = row_start+num_rows

		# Loop through each column and add series
		charts = []
		max_plots = len(color_map) - 1

		plots = self.split_plots(max_plots, graph_groups)

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

		return None

	def graph_grouping(self, df, group_by=constants.Results.chart_grouping, startcol=0):
		"""
			Determines sizes for grouping of graphs together
			CAVEAT:  Assumes that the DataFrame order matches with the order in Excel
		:param pd.DataFrame df: Dataframe to calculate grouping for
		:param tuple group_by: (optional = constants.Results.graph_grouping) = Levels to group by
		:param int startcol: (optional=0) Column which graphs start from to be added to the dataframe columns
		:return dict col_nums:  {Name of graph: Relative column numbers for results
		"""
		# Determine number of results in each group so that it can determine whether multiple graphs are needed
		df_groups = list(df.groupby(axis=1, level=group_by).size())

		# If only single plot on each graph then no need to separate at this level so go up 1 level
		if max(df_groups) == 1 and len(group_by)>1:
			self.logger.debug('Only single value for each entry so no need to split across multiple graphs')
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

	def split_plots(self, max_plots, graph_groups):
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

class PreviousResultsExport:
	""" Used for importing the setings and previously exported results """
	def __init__(self, pth):
		"""

		:param str pth:  path that will contain the input files
		"""

		self.logger = constants.logger

		self.search_pth = pth
		self.logger.debug('Processing results saved in: {}'.format(self.search_pth))

		# Get inputs used
		self.inputs = self.get_input_values()

		# Get a single DataFrame for all results
		self.df = self.import_all_results(study_type=constants.Results.study_fs)
		self.logger.debug('All results for folder: {} imported'.format(self.search_pth))

	def get_input_values(self):
		"""
			Function to import the inputs file found in a folder (only a single file is expected to be found)
		:param str search_pth:  Directory which contains the Inputs
		:return file_io.StudyInputs processed_inputs:  Processed import
		"""
		# Obtain reference to workbook from target directory
		c = constants.StudyInputs
		logger = constants.logger

		# Find the inputs file for this folder
		list_of_input_files = glob.glob('{}\{}*{}'.format(self.search_pth, c.file_name, c.file_format))
		if len(list_of_input_files) == 0:
			logger.critical(
				(
					'No inputs file formatted as *{} found in the folder {}, please check an inputs file exists'
				).format(c.file_name, c.file_format, self.search_pth)
			)
			raise IOError('No Inputs file found')
		elif len(list_of_input_files) > 1:
			inputs_workbook = list_of_input_files[0]
			logger.warning(
				(
					'Multiple input files were found in the folder {} with the format {}*{} as follows: \n\t {}\n'
					'The following file was assumed to be the correct one: {}'
				).format(self.search_pth, c.file_name, c.file_format, '\n\t'.join(list_of_input_files), inputs_workbook)
			)
		else:
			inputs_workbook = list_of_input_files[0]

		# Process the imported workbook into (gui_mode prevents the creation of a folder for the exports)
		processed_inputs = StudyInputs(inputs_workbook, gui_mode=True)
		# Set export folder = this folder
		processed_inputs.settings.export_folder = self.search_pth
		logger.info('Inputs from file: {} extracted'.format(inputs_workbook))
		return processed_inputs

	def import_all_results(self, study_type='FS'):
		"""
			Function to import all results into a single DataFrame
		:param str study_type: (Optional='FS') - Leading characters to use in search string
		:return pd.DataFrame single_df:  Combined imported files into single DataFrame
		"""
		self.study_type = study_type

		# Get list of all files in folder for frequency scan
		files = glob.glob('{}\{}*.csv'.format(self.search_pth, study_type))
		no_files = len(files)
		self.logger.debug('Importing {} results files in directory: {}'.format(no_files, self.search_pth))

		# Import each results file and combine into a single dataframe
		dfs = []
		for i, file in enumerate(files):
			df = self.process_file(pth=file)
			dfs.append(df)
			self.logger.info(' - \t {}/{} Results file: {} imported'.format(i+1, no_files, os.path.basename(file)))

		if len(dfs) != no_files:
			self.logger.error(
				(
					'There was an issue in the file import and not all were imported.\n '
					'Only {} of {} files were imported\n'
					'However, the script will continue until something critical occurs'
				)
					.format(len(dfs), no_files)
			)

		single_df = pd.concat(dfs, axis=1)
		self.logger.debug(
			'Single dataset for all results in folder:  {}'.format(self.search_pth)
		)
		return single_df

	def process_file(self, pth):
		"""
			# Process the imported results file into a dataframe with the relevant multi-index
		:param str pth:  Full path to results that need importing
		:return pd.DataFrame _df:  Return data frame processed ready for exporting to Excel in format
		"""
		c = constants.Results
		idx = pd.IndexSlice

		# Import dataframe
		df = pd.read_csv(pth, header=[0, 1])

		# set index based on frequency
		df.index = df.loc[:, idx[:, constants.PowerFactory.pf_freq]].squeeze()

		# Get from results file:
		#	Node / Mutual Name
		#	Variable type
		filename = os.path.basename(pth)

		# Process the file name to understand the study case being considered
		sc_name, cont_name = self.process_file_name(file_name=filename)

		# Create full name for study
		full_name = '{}_{}'.format(sc_name, cont_name)

		columns = list(zip(*df.columns.tolist()))
		# To manually deal with renaming of mutual impedance values
		var_names = columns[0]
		var_types = columns[1]

		# Mutual impedance dataframe
		df_mutual = pd.DataFrame().reindex_like(df)
		df_mutual = df_mutual.drop(df.columns, axis=1)
		new_var_names1 = []
		new_var_names2 = []
		new_ref_terminals1 = []
		new_ref_terminals2 = []
		new_var_types1 = []
		new_var_types2 = []
		for i, var in enumerate(var_names):
			var_names, ref_terms = self.extract_var_name(var_name=var)
			var_type = self.extract_var_type(var_types[i])

			# Mutual impedance data
			if type(var_names) is tuple:
				# If mutual impedance data is returned then need to duplicate dataframe to create data in the other direction
				new_var_names1.append(var_names[0])
				new_var_names2.append(var_names[1])
				new_ref_terminals1.append(ref_terms[0])
				new_ref_terminals2.append(ref_terms[1])
				new_var_types1.append(var_type)
				new_var_types2.append(var_type)

				# Create a copy of the data
				df_mutual[df.columns[i]] = df[df.columns[i]]
			else:
				# If no mutual impedance data then just extract variable names in lists
				new_var_names1.append(var_names)
				new_ref_terminals1.append(ref_terms)
				new_var_types1.append(var_type)

		# Produce new multi-index containing new headers
		col_headers1 = [(ref_terminal, var_name, sc_name, cont_name, full_name, var_type)
						for ref_terminal, var_name, var_type in zip(new_ref_terminals1, new_var_names1, new_var_types1)]
		col_headers2 = [(ref_terminal, var_name, sc_name, cont_name, full_name, var_type)
						for ref_terminal, var_name, var_type in zip(new_ref_terminals2, new_var_names2, new_var_types2)]

		# Names for headers
		names = (c.lbl_Reference_Terminal,
				 c.lbl_Terminal,
				 c.lbl_StudyCase,
				 c.lbl_Contingency,
				 c.lbl_FullName,
				 c.lbl_Result)

		# Produce multi-index and assign to DataFrames
		columns1 = pd.MultiIndex.from_tuples(tuples=col_headers1, names=names)
		columns2 = pd.MultiIndex.from_tuples(tuples=col_headers2, names=names)

		# Replace previous multi-index with new
		df.columns = columns1
		df_mutual.columns = columns2

		# Combine dataframes into one and return
		df = pd.concat([df, df_mutual], axis=1)

		# Obtain nominal voltage for each terminal
		idx = pd.IndexSlice
		# Add in row to contain the nominal voltage
		df.loc[c.idx_nom_voltage] = str()

		dict_nom_voltage = dict()

		# Find the nominal voltage for each terminal (if exists)
		for term, df_sub in df.groupby(axis=1, level=c.lbl_Reference_Terminal):
			idx_filter = idx[:,:,:,:,:,'e:uknom']
			try:
				# Obtain nominal voltage and then set row values appropriately to include in results
				nom_voltage = df.loc[:,idx_filter].iloc[0,0]
				dict_nom_voltage[term] = nom_voltage
			except KeyError:
				pass

		# Check for any duplicated multi-index entries (typically contingencies) and rename
		to_keep = 'first'
		if any(df.columns.duplicated(keep=to_keep)):
			self.logger.debug('Processing duplicated results in the results file: {}'.format(pth))
			# Get duplicated and non-duplicated into separate DataFrames
			duplicated_entries = df.loc[:, df.columns.duplicated(keep=to_keep)]
			non_duplicated_entries = df.loc[:, ~df.columns.duplicated(keep=to_keep)]

			# Rename duplicated_entries and report to user (contingency is the entry that is renamed)
			# Only allows for a single duplicated entry
			if any(duplicated_entries.columns.duplicated()):
				self.logger.critical(
					(
						'Unexpected error when trying to deal with duplicate columns for processing of the results '
						'file: {}'
					).format(pth)
				)
				raise IOError('Multiple duplicated entries')

			# Produce new column labels for each contingency
			terminals = set(duplicated_entries.columns.get_level_values(level=c.lbl_Reference_Terminal))
			msg = (
				(
					'During processing of the results file: {} some duplicated entries have been for the '
					'following terminals: \n'
				).format(pth)
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
			self.logger.warning(msg)

			# Combine the two DataFrames back into a single DataFrame
			df = pd.concat([duplicated_entries, non_duplicated_entries], axis=1)

		# Update the DataFrame to include the nominal voltage for every terminal
		for term, nom_voltage in dict_nom_voltage.items():
			df.loc[c.idx_nom_voltage, idx[term,:,:,:,:,:]] = nom_voltage

		# Add the new nominal voltage into the index data
		df = df.T.set_index(c.idx_nom_voltage, append=True).T
		df = df.reorder_levels([c.lbl_Reference_Terminal, c.lbl_Terminal, c.idx_nom_voltage,
								  c.lbl_StudyCase, c.lbl_Contingency,
								  c.lbl_FullName, c.lbl_Result], axis=1)

		return df

	def process_file_name(self, file_name):
		"""
			Splits up the file name to identify the study type, case and contingency
		:param str file_name:  Existing file name
		:return list components: [study_type, study_case, contingency) where file_name remaining
		"""
		c = constants.Results
		sc_name = ''
		cont_name = ''

		# Remove study_type and extension from filename
		file_name.replace('.csv', '')
		file_name.replace('{}{}'.format(self.study_type, c.joiner), '')

		# Find which study case is shown
		for sc in self.inputs.cases.index:
			if sc in file_name:
				sc_name = sc
				file_name = file_name.replace('{}{}'.format(sc_name, c.joiner), '')
				break

		# Find which contingnecy is considered
		# TODO: Check that _ symbol not used in studycase or contingency name
		for cont in self.inputs.contingencies.keys():
			if cont in file_name:
				cont_name = cont
				break

		return sc_name, cont_name

	def extract_var_name(self, var_name):
		"""
			Function extracts the variable name from the list
		:param str var_name: Name to extract relevant component from
		:return str var_name: Shortened name to use for processing (var1 = Substation or Mutual, var2 = Terminal)
		"""
		# Variable declarations
		c = constants.PowerFactory
		var_sub = False
		var_term = False
		ref_terminal = ''

		# Separate PowerFactory path into individual entries
		vars_list = var_name.split('\\')
		terminal_names = [x.name for x in self.inputs.terminals.values()]
		substations = [x.substation for x in self.inputs.terminals.values()]

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
					self.logger.error(
						(
							'Not completely sure if the reference terminals for mutual impedance {} are correct. In '
							'determining the reference terminals it is assumed that the mutual impedance is named '
							'as "Terminal 1_Terminal 2" but this seems not to be the case here.\n The following '
							'terminals have been identified: \n'
							'{}\n{}'
						).format(var_name, ref_terminal[0], ref_terminal[1])
					)
				# Two mutual names returns in lists for each direction and each ref_terminal
				# Variable names as lists in reverse order
				var_name = ('_'.join(ref_terminal),
							'_'.join(ref_terminal[::-1]))
				# Mutual name found so exit for loop
				break
			elif '.{}'.format(c.pf_substation) in var:
				var_sub = var
			elif '.{}'.format(c.pf_terminal) in var:
				var_term = var
			elif '.{}'.format(c.pf_results) in var:
				# This correlates to the frequency data but that has already been provided and so can
				# be deleted from the results
				var_name = constants.Results.lbl_to_delete
				ref_terminal = constants.Results.lbl_to_delete
				break

		# Lookup terminal name from input spreadsheet
		if ref_terminal == '':
			for term in self.inputs.terminals.values():
				# Find matching substation and terminal
				if term.substation == var_sub and term.terminal == var_term:
					ref_terminal = term.name
					break

			if ref_terminal == '':
				self.logger.critical(
					(
						'Substation {} in the results does not appear in the inputs list of substations with the '
						'associated terminal {}.  List of inputs are:\n\t{}'
					).format(var_sub, '\n\t'.join([term.name for term in self.inputs.terminals]))
				)
				raise ValueError(
					(
						'Substation {} not found in inputs {} but has appeared in results'
					).format(var_sub, self.search_pth)
				)

		return var_name, ref_terminal

	def extract_var_type(self, var_type):
		"""
			Function extracts the variable type by splitting at the first space
		:param str var_type: Typically provided in the format 'c:R_12 in Ohms'
		:return str (var_type): Shortened name to use for processing
		"""
		var_extract = var_type.split(' ')[0]

		# Raises an exception if poor data inputs given
		if var_extract not in constants.StudyInputs.all_variable_types:
			raise IOError('The variable extracted {} from {} is not one of the input types {}'
						  .format(var_extract, var_type, constants.StudyInputs.all_variable_types))
		return var_extract

class StudySettings:
	"""
		Class contains the processing of each of the DataFrame items passed as part of the
	"""
	def __init__(self, sht=constants.StudyInputs.study_settings, wkbk=None, pth_file=None, gui_mode=False):
		"""
			Process the worksheet to extract the relevant StudyCase details
		:param str sht:  (optional) Name of worksheet to use
		:param pd.ExcelFile wkbk:  (optional) Handle to workbook
		:param str pth_file: (optional) Handle to workbook
		:param bool gui_mode: (optional) When set to True this will prevent the creation of a new
							folder since the user will provide the folder when running the studies
		"""
		# Constants used as part of this
		self.export_folder = str()
		self.results_name = str()
		self.pre_case_check = bool()
		self.delete_created_folders = bool()
		self.export_to_excel = bool()
		self.export_rx = bool()
		self.export_mutual = bool()
		self.include_intact = bool()

		self.c = constants.StudySettings
		self.logger = constants.logger

		# Sheet name
		self.sht = sht

		# GUI Mode - Prevents creation of folders that are defined later
		self.gui_mode = gui_mode

		# Import workbook as dataframe
		if wkbk is None:
			if pth_file:
				wkbk = pd.ExcelFile(pth_file)
				self.pth = pth_file
			else:
				raise IOError('No workbook or path to file provided')
		else:
			# Get workbook path in case path has not been provided
			self.pth = wkbk.io

		# Import Study settings into a DataFrame and process
		self.df = pd.read_excel(
			wkbk, sheet_name=self.sht, index_col=0, usecols=(0, 1), skiprows=4, header=None, squeeze=True
		)
		self.process_inputs()

	def get_vars_to_export(self):
		"""
			Determines the variables that will be exported from PowerFactory and they will be exported in this order
		:return list pf_vars:  Returns list of variables in the format they are defined in PowerFactory
		"""
		c = constants.PowerFactory
		pf_vars = [c.pf_z1, ]

		# The order variables are added here determines the order they appear in the export
		# If self impedance data should be exported
		if self.export_rx:
			pf_vars.append(c.pf_r1)
			pf_vars.append(c.pf_x1)

		# If mutual impedance data should be exported
		if self.export_mutual:
			# If RX data should be exported
			if self.export_rx:
				pf_vars.append(c.pf_r12)
				pf_vars.append(c.pf_x12)
			pf_vars.append(c.pf_z12)

		return pf_vars

	def process_inputs(self):
		""" Process all of the inputs into attributes of the class """
		# Process results_folder
		if not self.gui_mode:
			self.export_folder = self.process_export_folder()
		else:
			self.logger.debug('Running in GUI mode and therefore export_folder is defined later')
		# self.results_name = self.process_result_name()
		# self.pf_network_elm = self.process_net_elements()

		self.pre_case_check = self.process_booleans(key=self.c.pre_case_check)
		self.delete_created_folders = self.process_booleans(key=self.c.delete_created_folders)
		self.export_to_excel = self.process_booleans(key=self.c.export_to_excel)
		self.export_rx = self.process_booleans(key=self.c.export_rx)
		self.export_mutual = self.process_booleans(key=self.c.export_mutual)
		self.include_intact = self.process_booleans(key=self.c.include_intact)

		# Sanity check for Boolean values
		self.boolean_sanity_check()

		return None

	def process_export_folder(self, def_value=os.path.join(os.path.dirname(__file__), '..')):
		"""
			Process the export folder and if a blank value is provided then use the default value
		:param str def_value:  Default path to use
		:return str folder:
		"""
		# Get folder from DataFrame, if empty or invalid path then use default folder
		folder = self.df.loc[self.c.export_folder]
		# Normalise path
		def_value = os.path.join(os.path.normpath(def_value), constants.uid)

		if not folder:
			# If no folder provided then use default value
			folder = os.path.normpath(def_value)
		elif not os.path.isdir(os.path.dirname(folder)):
			# If folder provided but not a valid directory for display warning message to user and use default directory
			self.logger.warning((
				'The parent directory for the results export path: {} does not exist and therefore the default directory '
				'of {} has been used instead'
			).format(os.path.dirname(folder), def_value))
			folder = def_value

		# Check if target directory exists and if not create
		if not os.path.isdir(folder):
			os.mkdir(folder)
			self.logger.info('The directory for the results {} does not exist but has been created'.format(folder))

		return folder

	def add_folder(self, pth_results_file):
		"""
			Folder creates a results file
		:param str pth_results_file:
		:return:
		"""
		delay_counter = 5

		# Get the full path to the results file
		pth, file_name = os.path.split(pth_results_file)
		foldername, _ = os.path.splitext(file_name)

		target_folder = os.path.join(pth, foldername)

		if os.path.exists(target_folder):
			self.logger.warning(
				(
					'The target folder for the results already exists, this will now be deleted including all of '
					'its contents.  The script will wait {} seconds before doing this during which time you can stop it'
				)
			)
			for x in range(delay_counter):
				self.logger.warning('Waiting {} / {} seconds'.format(x, delay_counter))
				time.sleep(1)
			shutil.rmtree(target_folder)

		self.logger.debug('Creating folder {} for the results'.format(target_folder))

		os.mkdir(target_folder)
		self.export_folder = target_folder

		return None


	# def process_result_name(self, def_value=constants.StudySettings.def_results_name):
	# 	"""
	# 		Processes the results file name
	# 	:param str def_value:  (optional) Default value to use
	# 	:return str results_name:
	# 	"""
	# 	results_name = self.df.loc[self.c.results_name]
	#
	# 	if not results_name:
	# 		# If no value provided then use default value
	# 		self.logger.warning((
	# 			'No value provided in the Input Settings for the results name and so the default value of {} will be '
	# 			'used instead'
	# 		).format(def_value)
	# 		)
	# 		results_name = def_value
	#
	# 	# Add study_time to end of results name
	# 	results_name = '{}_{}{}'.format(results_name, self.uid, constants.Extensions.excel)
	#
	# 	return results_name

	def process_booleans(self, key):
		"""
			Function imports the relevant boolean value and confirms it is either True / False, if empty then just
			raises warning message to the user
		:param str key:
		:return bool value:
		"""
		# Get folder from DataFrame, if empty or invalid path then use default folder
		value = self.df.loc[key]

		if value == '':
			value = False
			self.logger.warning(
				(
					'No value has been provided for key item <{}> in worksheet <{}> of the input file <{}> and so '
					'{} will be assumed'
				).format(
					key, self.sht, self.pth, value
				)
			)
		else:
			# Ensure value is a suitable Boolean
			value = bool(value)

		return value

	def boolean_sanity_check(self):
		"""
			Function to check if any of the input Booleans have not been set which require something to be set for
			results to be of any use.

			Throws an error if nothing suitable / logs a warning message

		:return None:
		"""

		if not self.export_to_excel and self.delete_created_folders:
			self.logger.critical((
				'Your input settings for {} = {} means that no results are exported.  However, you are also deleting '
				'all the created results with the command {} = {}.  This is probably not what you meant to happen so '
				'I am stopping here for you to correct your inputs!').format(
				self.c.export_to_excel, self.export_to_excel, self.c.delete_created_folders, self.delete_created_folders)
			)
			raise ValueError('Your input settings mean nothing would be produced')

		if not self.pre_case_check:
			self.logger.warning((
				'You have opted not to run a pre-case check with the input {} = {}, this means that if there are any '
				'issues with an individual contingency the entire study set may fail'
			).format(self.c.pre_case_check, self.pre_case_check))

		return None

class StudyInputs:
	"""
		Class used to import the Settings from the Input Spreadsheet and convert into a format usable elsewhere
	"""
	def __init__(self, pth_file=None, gui_mode=False):
		"""
			Initialises the settings based on the Study Settings spreadsheet
		:param str pth_file:  Path to input settings file
		:param bool gui_mode: (optional) when set to True this will creating the export folder since that
							will be processed later
		"""
		# General constants
		self.pth = pth_file
		self.filename = os.path.basename(pth_file)
		self.logger = constants.logger

		# TODO: Adjust to update if an Import Error occurs
		self.error = False

		self.logger.info('Importing settings from file: {}'.format(self.pth))

		with pd.ExcelFile(io=self.pth) as wkbk:
			# Import StudySettings
			self.settings = StudySettings(wkbk=wkbk, gui_mode=gui_mode)
			self.cases = self.process_study_cases(wkbk=wkbk)
			self.contingency_cmd, self.contingencies = self.process_contingencies(wkbk=wkbk)
			self.terminals = self.process_terminals(wkbk=wkbk)
			self.lf_settings = self.process_lf_settings(wkbk=wkbk)
			self.fs_settings = self.process_fs_settings(wkbk=wkbk)

		if not self.error:
			self.logger.info('Importing settings from file: {} completed'.format(self.pth))
		else:
			self.logger.warning('Error during import of settings from file: {}'.format(self.pth))

	def load_workbook(self, pth_file=None):
		"""
			Function to load the workbook and return a handle to it
		:param str pth_file: (optional) File path to workbook
		:return pd.ExcelFile wkbk: Handle to workbook
		"""
		if pth_file:
			# Load the workbook using pd.ExcelFile and return the reference to the workbook
			wkbk = pd.ExcelFile(pth_file)
			self.pth = pth_file
		else:
			raise IOError('No workbook or path to file provided')

		return wkbk

	def copy_inputs_file(self):
		"""
			Function copies the inputs file to the desired results folder and applies the appropriate header
		:return:
		"""

		src = self.pth

		# Produce new name for inputs file
		file_name = os.path.basename(src)
		dest = os.path.join(
			self.settings.export_folder, '{}_{}{}'.format(
				constants.StudyInputs.file_name,
				constants.uid,
				file_name
			)
		)

		# Copy to destination
		self.logger.debug('Saving inputs file {} in to export folder: {}'.format(src, dest))
		shutil.copyfile(src=src, dst=dest)

		return None

	def process_study_cases(self, sht=constants.StudyInputs.study_cases, wkbk=None, pth_file=None):
		"""
			Function imports the DataFrame of study cases and then separates each one into it's own study case.  These
			are then returned as a dictionary with the name being used as the key.

			These inputs are based on the Scenarios detailed in the Inputs spreadsheet

			:param str sht:  (optional) Name of worksheet to use
			:param pd.ExcelFile wkbk:  (optional) Handle to workbook
			:param str pth_file: (optional) Handle to workbook
		:return pd.DataFrame df:  Returns a DataFrame of the study cases so can group by columns
		"""

		# Import workbook as dataframe
		if wkbk is None:
			wkbk = self.load_workbook(pth_file=pth_file)

		# Import Study settings into a DataFrame and process, do not need to worry about unique columns since done by
		# position
		df = pd.read_excel(wkbk, sheet_name=sht, skiprows=3, header=0, usecols=(0, 1, 2, 3))
		cols = constants.StudySettings.studycase_columns
		df.columns = cols

		# Process df names to confirm unique
		name_key = cols[0]
		df, updated = update_duplicates(key=name_key, df=df)

		# Make index of study_cases name after duplicates removed
		df.set_index(constants.StudySettings.name, inplace=True, drop=False)

		if updated:
			self.logger.warning(
				(
					'Duplicated names for StudyCases provided in the column {} of worksheet <{}> and so some have been '
					'renamed so the new list of names is:\n\t{}'
				).format(name_key, sht, '\n\t'.join(df.loc[:, name_key].values))
			)

		# # Iterate through each DataFrame and create a study case instance and OrderedDict used to ensure no change in
		# # order
		# study_cases = collections.OrderedDict()
		# for key, item in df.iterrows():
		# 	new_case = StudyCaseDetails(list_of_parameters=item.values)
		# 	study_cases[new_case.name] = new_case

		return df

	def process_contingencies(self, sht=constants.StudyInputs.contingencies, wkbk=None, pth_file=None):
		"""
			Function imports the DataFrame of contingencies and then each into its own contingency under a single
			dictionary key.  Each dictionary item contains the relevant outages to be taken in the form
			ContingencyDetails.

			Also returns a string detailing the name of a Contingency file if one is provided

		:param str sht:  (optional) Name of worksheet to use
		:param pd.ExcelFile wkbk:  (optional) Handle to workbook
		:param str pth_file: (optional) File path to workbook
		:return (str, dict) (contingency_cmd, contingencies):  Returns both:
																The command for all contingencies
																A dictionary with each of the outages
		"""

		# Import workbook as dataframe
		if wkbk is None:
			wkbk = self.load_workbook(pth_file=pth_file)

		# Get details of existing contingency command and if exists set the reference command
		# Squeeze command to ensure converted to a series since only a single row is being imported
		df = pd.read_excel(
			wkbk, sheet_name=sht, skiprows=2, nrows=1, usecols=(0,1,2,3), index_col=None, header=None
		).squeeze(axis=0)
		cmd = df.iloc[3]
		# Valid command is a string whereas non-valid commands will be either False or an empty string
		if cmd:
			contingency_cmd = cmd
		else:
			contingency_cmd = str()


		# Import rest of details of contingencies into a DataFrame and process, do not need to worry about unique
		# columns since done by position
		df = pd.read_excel(wkbk, sheet_name=sht, skiprows=3, header=(0, 1))
		cols = df.columns

		# Process df names to confirm unique
		name_key = cols[0]
		df, updated = update_duplicates(key=name_key, df=df)

		# Make index of contingencies name after duplicates removed
		df.set_index(name_key, inplace=True, drop=False)


		if updated:
			self.logger.warning(
				(
					'Duplicated names for Contingencies provided in the column {} of worksheet <{}> and so some have been '
					'renamed and the new list of names is:\n\t{}'
				).format(name_key, sht, '\n\t'.join(df.loc[:, name_key].values))
			)

		# Iterate through each DataFrame and create a study case instance and OrderedDict used to ensure no change in order
		contingencies = collections.OrderedDict()
		for key, item in df.iterrows():
			cont = ContingencyDetails(list_of_parameters=item.values)
			contingencies[cont.name] = cont

		return contingency_cmd, contingencies

	def process_terminals(self, sht=constants.StudyInputs.terminals, wkbk=None, pth_file=None):
		"""
			Function imports the DataFrame of terminals.
			These are then returned as a dictionary with the name being used as the key.

			These inputs are based on the Scenarios detailed in the Inputs spreadsheet

		:param str sht:  (optional) Name of worksheet to use
		:param pd.ExcelFile wkbk:  (optional) Handle to workbook
		:param str pth_file: (optional) File path to workbook
		:return dict terminals: type: TerminalDetails
		"""

		# Import workbook as dataframe
		if wkbk is None:
			wkbk = self.load_workbook(pth_file=pth_file)

		# Import Contingencies into a DataFrame and process, do not need to worry about unique columns since done by
		# position
		df = pd.read_excel(wkbk, sheet_name=sht, skiprows=3, header=0)
		cols = df.columns

		# Process df names to confirm unique
		name_key = cols[0]
		df, updated = update_duplicates(key=name_key, df=df)

		# Make index of terminals name after duplicates removed
		df.set_index(name_key, inplace=True, drop=False)

		if updated:
			self.logger.warning(
				(
					'Duplicated names for Terminals provided in the column {} of worksheet <{}> and so some have been '
					'renamed and the new list of names is:\n\t{}'
				).format(name_key, sht, '\n\t'.join(df.loc[:, name_key].values))
			)

		# Iterate through each DataFrame and create a study case instance and OrderedDict used to ensure no change in order
		terminals = collections.OrderedDict()
		for key, item in df.iterrows():
			term = TerminalDetails(list_of_parameters=item.values)
			terminals[term.name] = term

		return terminals

	def process_lf_settings(self, sht=constants.StudyInputs.lf_settings, wkbk=None, pth_file=None):
		"""
			Process the provided load flow settings
		:param str sht:  (optional) Name of worksheet to use
		:param pd.ExcelFile wkbk:  (optional) Handle to workbook
		:param str pth_file: (optional) File path to workbook
		:return LFSettings lf_settings:  Returns reference to instance of load_flow settings
		"""
		# Import workbook as dataframe
		if wkbk is None:
			wkbk = self.load_workbook(pth_file=pth_file)

		# Import Load flow settings into a DataFrame and process
		df = pd.read_excel(
			wkbk, sheet_name=sht, usecols=(3, ), skiprows=3, header=None, squeeze=True
		)

		lf_settings = LFSettings(existing_command=df.iloc[0], detailed_settings=df.iloc[1:])
		return lf_settings

	def process_fs_settings(self, sht=constants.StudyInputs.fs_settings, wkbk=None, pth_file=None):
		"""
			Process the provided frequency sweep settings
		:param str sht:  (optional) Name of worksheet to use
		:param pd.ExcelFile wkbk:  (optional) Handle to workbook
		:param str pth_file: (optional) File path to workbook
		:return LFSettings lf_settings:  Returns reference to instance of load_flow settings
		"""
		# Import workbook as dataframe
		if wkbk is None:
			wkbk = self.load_workbook(pth_file=pth_file)

		# Import Load flow settings into a DataFrame and process
		df = pd.read_excel(
			wkbk, sheet_name=sht, usecols=(3, ), skiprows=3, header=None, squeeze=True
		)

		settings = FSSettings(existing_command=df.iloc[0], detailed_settings=df.iloc[1:])
		return settings

class StudyCaseDetails:
	def __init__(self, list_of_parameters):
		"""
			Single row of study case parameters imported from spreadsheet are defined into a class for easy lookup
		:param list list_of_parameters:  Single row of inputs
		"""
		self.name = list_of_parameters[0]
		self.project = list_of_parameters[1]
		self.study_case = list_of_parameters[2]
		self.scenario = list_of_parameters[3]

class ContingencyDetails:
	def __init__(self, list_of_parameters):
		"""
			Single row of contingency parameters imported from spreadsheet are defined into a class for easy lookup
		:param list list_of_parameters:  Single row of inputs
		"""
		# Status flag set to True if cannot be found, contingency fails, etc.
		self.not_included = False

		self.name = list_of_parameters[0]
		self.couplers = []
		for substation, breaker, status in zip(*[iter(list_of_parameters[1:])]*3):
			if substation != '' and breaker != '' and str(breaker) != 'nan':
				new_coupler = CouplerDetails(substation, breaker, status)
				self.couplers.append(new_coupler)

		# If contingency has been defined then needs to be included in results
		# # Check if this contingency relates to the intact system in which case it will be skipped
		if len(self.couplers) == 0 or self.name == constants.StudyInputs.base_case:
			self.skip = True
		else:
			self.skip = False

class CouplerDetails:
	def __init__(self, substation, breaker, status):
		"""
			Define the substations, breaker and status
		:param str substation:  Name of substation
		:param str breaker: Name of breaker
		:param bool status: Status breaker is changed to
		"""
		# Check if substation already has substation type ending and if not add it
		if not str(substation).endswith(constants.PowerFactory.pf_substation):
			substation = '{}.{}'.format(substation, constants.PowerFactory.pf_substation)

		# Check if breaker ends in the correct format
		if not str(breaker).endswith(constants.PowerFactory.pf_coupler):
			breaker = '{}.{}'.format(breaker, constants.PowerFactory.pf_coupler)

		# Confirm that status for operation is either true or false
		if status == constants.StudyInputs.switch_open:
			status = False
		elif status == constants.StudyInputs.switch_close:
			status = True
		else:
			logger = constants.logger
			logger.warning(
				(
					'The breaker <{}> associated with substation <{}> has a value of {} which is not expected value of '
					'{} or {}.  The operation is assumed to be {} for this study but the user should check '
					'that is what they intended'
				).format(
					breaker, substation, status,
					constants.StudyInputs.switch_open, constants.StudyInputs.switch_close,
					constants.StudyInputs.switch_open
				)
			)
			status = False

		self.substation = substation
		self.breaker = breaker
		self.status = status

class TerminalDetails:
	"""
		Details for each terminal that data is required for from processing
	"""
	def __init__(self, name=str(), substation=str(), terminal=str(), include_mutual=True, list_of_parameters=list()):
		"""
			Process each terminal
		:param list list_of_parameters: (optional=none)
		:param str name:  Input name to use
		:param str substation:  Name of substation within which terminal is contained
		:param str terminal:   Name of terminal in substation
		:param bool include_mutual:  (optional=True) - If mutual impedance data is not required for this terminal then
			set to False
		"""

		if len(list_of_parameters) > 0:
			name = str(list_of_parameters[0])
			substation = str(list_of_parameters[1])
			terminal = str(list_of_parameters[2])
			include_mutual = bool(list_of_parameters[3])

		# Add in the relevant endings if they do not already exist
		c = constants.PowerFactory
		if not substation.endswith(c.pf_substation):
			substation = '{}.{}'.format(substation, c.pf_substation)

		if not terminal.endswith(c.pf_terminal):
			terminal = '{}.{}'.format(terminal, c.pf_terminal)

		self.name = name
		self.substation = substation
		self.terminal = terminal
		self.include_mutual = include_mutual

		# The following are populated once looked for in the specific PowerFactory project
		self.pf_handle = None
		self.found = None

class LFSettings:
	def __init__(self, existing_command, detailed_settings):
		"""
			Initialise variables
		:param str existing_command:  Reference to an existing command where it already exists
		:param list detailed_settings:  Settings to be used where existing command does not exist
		"""
		self.logger = constants.logger

		# Add the Load Flow command to string
		if existing_command:
			if not existing_command.endswith('.{}'.format(constants.PowerFactory.ldf_command)):
				existing_command = '{}.{}'.format(existing_command, constants.PowerFactory.ldf_command)
		else:
			existing_command = None
		self.cmd = existing_command

		# Basic
		self.iopt_net = int()  # Calculation method (0 Balanced AC, 1 Unbalanced AC, DC)
		self.iopt_at = int()  # Automatic Tap Adjustment
		self.iopt_asht = int()  # Automatic Shunt Adjustment

		# Added in Automatic Tapping of PSTs but for backwards compatibility will ensure can work when less than 1
		self.iPST_at = int()  # Automatic Tap Adjustment of Phase Shifters

		self.iopt_lim = int()  # Consider Reactive Power Limits
		self.iopt_limScale = int()  # Consider Reactive Power Limits Scaling Factor
		self.iopt_tem = int()  # Temperature Dependency: Line Cable Resistances (0 ...at 20C, 1 at Maximum Operational Temperature)
		self.iopt_pq = int()  # Consider Voltage Dependency of Loads
		self.iopt_fls = int()  # Feeder Load Scaling
		self.iopt_sim = int()  # Consider Coincidence of Low-Voltage Loads
		self.scPnight = int()  # Scaling Factor for Night Storage Heaters

		# Active Power Control
		self.iopt_apdist = int()  # Active Power Control (0 as Dispatched, 1 According to Secondary Control,
		# 2 According to Primary Control, 3 According to Inertia)
		self.iopt_plim = int()  # Consider Active Power Limits
		self.iPbalancing = int()  # (0 Ref Machine, 1 Load, Static Gen, Dist slack by loads, Dist slack by Sync,

		# Get DataObject handle for reference busbar
		self.substation = str()
		self.terminal = str()

		self.phiini = int() # Angle

		# Advanced Options
		self.i_power = int()  # Load Flow Method ( NR Current, 1 NR (Power Eqn Classic)
		self.iopt_notopo = int()  # No Topology Rebuild
		self.iopt_noinit = int()  # No initialisation
		self.utr_init = int()  # Consideration of transformer winding ratio
		self.maxPhaseShift = int()  # Max Transformer Phase Shift
		self.itapopt = int()  # Tap Adjustment ( 0 Direct, 1 Step)
		self.krelax = int()  # Min Controller Relaxation Factor

		self.iopt_stamode = int()  # Station Controller (0 Standard, 1 Gen HV, 2 Gen LV
		self.iopt_igntow = int()  # Modelling Method of Towers (0 With In/ Output signals, 1 ignore couplings, 2 equation in lines)
		self.initOPF = int()  # Use this load flow for initialisation of OPF
		self.zoneScale = int()  # Zone Scaling ( 0 Consider all loads, 1 Consider adjustable loads only)

		# Iteration Control
		self.itrlx = int()  # Max No Iterations for Newton-Raphson Iteration
		self.ictrlx = int()  # Max No Iterations for Outer Loop
		self.nsteps = int()  # Max No Iterations for Number of steps

		self.errlf = int()  # Max Acceptable Load Flow Error for Nodes
		self.erreq = int()  # Max Acceptable Load Flow Error for Model Equations
		self.iStepAdapt = int()  # Iteration Step Size ( 0 Automatic, 1 Fixed Relaxation)
		self.relax = int()  # If Fixed Relaxation factor
		self.iopt_lev = int()  # Automatic Model Adaptation for Convergence

		# Outputs
		self.iShowOutLoopMsg = int()  # Show 'outer Loop' Messages
		self.iopt_show = int()  # Show Convergence Progress Report
		self.num_conv = int()  # Number of reported buses/models per iteration
		self.iopt_check = int()  # Show verification report
		self.loadmax = int()  # Max Loading of Edge Element
		self.vlmin = int()  # Lower Limit of Allowed Voltage
		self.vlmax = int()  # Upper Limit of Allowed Voltage
		self.iopt_chctr = int()  # Check Control Conditions

		# Load Generation Scaling
		self.scLoadFac = int()  # Load Scaling Factor
		self.scGenFac = int()  # Generation Scaling Factor
		self.scMotFac = int()  # Motor Scaling Factor

		# Low Voltage Analysis
		self.Sfix = int()  # Fixed Load kVA
		self.cosfix = int()  # Power Factor of Fixed Load
		self.Svar = int()  # Max Power Per Customer kVA
		self.cosvar = int()  # Power Factor of Variable Part
		self.ginf = int()  # Coincidence Factor
		self.i_volt = int()  # Voltage Drop Analysis (0 Stochastic Evaluation, 1 Maximum Current Estimation)

		# Advanced Simulation Options
		self.iopt_prot = int()  # Consider Protection Devices ( 0 None, 1 all, 2 Main, 3 Backup)
		self.ign_comp = int()  # Ignore Composite Elements

		# Value set to True if error occurs when processing settings
		self.settings_error = False

		try:
			self.populate_data(load_flow_settings=list(detailed_settings))
		except (IndexError, AttributeError):
			if self.cmd:
				self.logger.warning(
					(
						'There were some missing settings in the load flow settings input but an existing command {} '
						'is being used instead. If this command is missing from the Study Cases then the script will '
						'fail'
					).format(self.cmd)
				)
				self.settings_error = True
			else:
				self.logger.warning(
					(
						'The load flow settings provided are incorrect / missing some values and no input for an existing '
						'load flow command (.{}) has been provided.  The study will just use the default load flow '
						'command associated with each study case but results may be inconsistent!'
					).format(constants.PowerFactory.ldf_command)
				)
				self.settings_error = True


	def populate_data(self, load_flow_settings):
		"""
			List of settings for the load flow from if using a manual settings file
		:param list load_flow_settings:
		"""

		# Loadflow settings
		# -------------------------------------------------------------------------------------
		# Create new object for the load flow on the base case so that existing settings are not overwritten

		# Get handle for load flow command from study case
		# Basic
		self.iopt_net = load_flow_settings[0]  # Calculation method (0 Balanced AC, 1 Unbalanced AC, DC)
		self.iopt_at = load_flow_settings[1]  # Automatic Tap Adjustment
		self.iopt_asht = load_flow_settings[2]  # Automatic Shunt Adjustment

		# Added in Automatic Tapping of PSTs with default values
		self.iPST_at = load_flow_settings[3]  # Automatic Tap Adjustment of Phase Shifters

		self.iopt_lim = load_flow_settings[4]  # Consider Reactive Power Limits
		self.iopt_limScale = load_flow_settings[5]  # Consider Reactive Power Limits Scaling Factor
		self.iopt_tem = load_flow_settings[6]  # Temperature Dependency: Line Cable Resistances (0 ...at 20C, 1 at Maximum Operational Temperature)
		self.iopt_pq = load_flow_settings[7]  # Consider Voltage Dependency of Loads
		self.iopt_fls = load_flow_settings[8]  # Feeder Load Scaling
		self.iopt_sim = load_flow_settings[9]  # Consider Coincidence of Low-Voltage Loads
		self.scPnight = load_flow_settings[10]  # Scaling Factor for Night Storage Heaters

		# Active Power Control
		self.iopt_apdist = load_flow_settings[11]  # Active Power Control (0 as Dispatched, 1 According to Secondary Control,
		# 2 According to Primary Control, 3 According Inertia)
		self.iopt_plim = load_flow_settings[12]  # Consider Active Power Limits
		self.iPbalancing = load_flow_settings[13]  # (0 Ref Machine, 1 Load, Static Gen, Dist slack by loads, Dist slack by Sync,

		# Get DataObject handle for reference busbar
		ref_machine = load_flow_settings[14]

		# To avoid error when passed empty string or empty cell which is imported as pd.na
		if not ref_machine or pd.isna(ref_machine):
			self.substation = None
			self.terminal = None
		else:
			net_folder_name, substation, terminal = ref_machine.split('\\')
			# Confirm that substation and terminal types exist in name
			if not substation.endswith(constants.PowerFactory.pf_substation):
				self.substation = '{}.{}'.format(substation, constants.PowerFactory.pf_substation)
			if not terminal.endswith(constants.PowerFactory.pf_terminal):
				self.terminal = '{}.{}'.format(terminal, constants.PowerFactory.pf_terminal)


		self.phiini = load_flow_settings[15]  # Angle

		# Advanced Options
		self.i_power = load_flow_settings[16]  # Load Flow Method ( NR Current, 1 NR (Power Eqn Classic)
		self.iopt_notopo = load_flow_settings[17]  # No Topology Rebuild
		self.iopt_noinit = load_flow_settings[18]  # No initialisation
		self.utr_init = load_flow_settings[19]  # Consideration of transformer winding ratio
		self.maxPhaseShift = load_flow_settings[20]  # Max Transformer Phase Shift
		self.itapopt = load_flow_settings[21]  # Tap Adjustment ( 0 Direct, 1 Step)
		self.krelax = load_flow_settings[22]  # Min Controller Relaxation Factor

		self.iopt_stamode = load_flow_settings[23]  # Station Controller (0 Standard, 1 Gen HV, 2 Gen LV
		self.iopt_igntow = load_flow_settings[24]  # Modelling Method of Towers (0 With In/ Output signals, 1 ignore couplings, 2 equation in lines)
		self.initOPF = load_flow_settings[25]  # Use this load flow for initialisation of OPF
		self.zoneScale = load_flow_settings[26]  # Zone Scaling ( 0 Consider all loads, 1 Consider adjustable loads only)

		# Iteration Control
		self.itrlx = load_flow_settings[27]  # Max No Iterations for Newton-Raphson Iteration
		self.ictrlx = load_flow_settings[28]  # Max No Iterations for Outer Loop
		self.nsteps = load_flow_settings[29]  # Max No Iterations for Number of steps

		self.errlf = load_flow_settings[30]  # Max Acceptable Load Flow Error for Nodes
		self.erreq = load_flow_settings[31]  # Max Acceptable Load Flow Error for Model Equations
		self.iStepAdapt = load_flow_settings[32]  # Iteration Step Size ( 0 Automatic, 1 Fixed Relaxation)
		self.relax = load_flow_settings[33]  # If Fixed Relaxation factor
		self.iopt_lev = load_flow_settings[34]  # Automatic Model Adaptation for Convergence

		# Outputs
		self.iShowOutLoopMsg = load_flow_settings[35]  # Show 'outer Loop' Messages
		self.iopt_show = load_flow_settings[36]  # Show Convergence Progress Report
		self.num_conv = load_flow_settings[37]  # Number of reported buses/models per iteration
		self.iopt_check = load_flow_settings[38]  # Show verification report
		self.loadmax = load_flow_settings[39]  # Max Loading of Edge Element
		self.vlmin = load_flow_settings[40]  # Lower Limit of Allowed Voltage
		self.vlmax = load_flow_settings[41]  # Upper Limit of Allowed Voltage
		self.iopt_chctr = load_flow_settings[42]  # Check Control Conditions

		# Load Generation Scaling
		self.scLoadFac = load_flow_settings[43]  # Load Scaling Factor
		self.scGenFac = load_flow_settings[44]  # Generation Scaling Factor
		self.scMotFac = load_flow_settings[45]  # Motor Scaling Factor

		# Low Voltage Analysis
		self.Sfix = load_flow_settings[46]  # Fixed Load kVA
		self.cosfix = load_flow_settings[47]  # Power Factor of Fixed Load
		self.Svar = load_flow_settings[48]  # Max Power Per Customer kVA
		self.cosvar = load_flow_settings[49]  # Power Factor of Variable Part
		self.ginf = load_flow_settings[50]  # Coincidence Factor
		self.i_volt = load_flow_settings[51]  # Voltage Drop Analysis (0 Stochastic Evaluation, 1 Maximum Current Estimation)

		# Advanced Simulation Options
		self.iopt_prot = load_flow_settings[52]  # Consider Protection Devices ( 0 None, 1 all, 2 Main, 3 Backup)
		self.ign_comp = load_flow_settings[53]  # Ignore Composite Elements

	def find_reference_terminal(self, app):
		"""
			Find and populate reference terminal for machine
		:param powerfactory.app app:
		:return None:
		"""
		# Confirm that a machine has actually been provided
		if self.substation is None or self.terminal is None:
			pf_term = None
			self.logger.debug('No reference machine provided in inputs')

		else:
			pf_sub = app.GetCalcRelevantObjects(self.substation)

			if len(pf_sub) == 0:
				pf_term = None
			else:
				pf_term = pf_sub[0].GetContents(self.terminal)
				if len(pf_term) == 0:
					pf_term = None

			if pf_term is None:
				self.logger.warning(
					(
						'Either the substation {} or terminal {} detailed as being the reference busbar cannot be '
						'found and therefore no reference busbar will be defined in the PowerFactory load flow'
					).format(self.substation, self.terminal)
				)

		return pf_term

class FSSettings:
	""" Class contains the frequency scan settings """
	def __init__(self, existing_command, detailed_settings):
		"""
			Initialise variables
		:param str existing_command:  Reference to an existing command where it already exists
		:param list detailed_settings:  Settings to be used where existing command does not exist
		"""
		self.logger = constants.logger

		# Add the Load Flow command to string
		if existing_command:
			if not existing_command.endswith('.{}'.format(constants.PowerFactory.fs_command)):
				existing_command = '{}.{}'.format(existing_command, constants.PowerFactory.fs_command)
		else:
			existing_command = None
		self.cmd = existing_command

		# Nominal frequency of system
		self.frnom = float()

		# Basic
		self.iopt_net = int() # Network representation (0 Balanced, 1 Unbalanced)
		self.fstart = float() # Impedance calculation start frequency
		self.fstop = float() # Impedance calculation stop frequency
		self.fstep = float() # Impedance calculation step size
		self.i_adapt = bool() # Automatic step size adaption (False = No, True = Yes)

		# Advanced
		self.errmax = float() # Setting for step size adaption maximum prediction error
		self.errinc = float() # Setting for step size minimum prediction error
		self.ninc = int() # Step size increase delay
		self.ioutall = bool() # Calculate R, X at output frequency for all nodes (False = No, True = Yes)

		# Value set to True if error occurs when processing settings
		self.settings_error = False

		try:
			self.populate_data(fs_settings=list(detailed_settings))
		except (IndexError, AttributeError):
			if self.cmd:
				self.logger.warning(
					(
						'There were some missing settings in the frequency sweep settings input but an existing command {} '
						'is being used instead. If this command is missing from the Study Cases then the script will '
						'fail'
					).format(self.cmd)
				)
				self.settings_error = True
			else:
				self.logger.warning(
					(
						'The frequency sweep settings provided are incorrect / missing some values and no input for an '
						'existing frequency sweep command (.{}) has been provided.  The study will just use the default '
						'frequency sweep command associated with each study case but results may be inconsistent!'
					).format(constants.PowerFactory.fs_command)
				)
				self.settings_error = True


	def populate_data(self, fs_settings):
		"""
			List of settings for the load flow from if using a manual settings file
		:param list fs_settings:
		"""

		# Nominal frequency of system
		self.frnom = fs_settings[0]

		# Basic
		self.iopt_net = int(fs_settings[1])
		self.fstart = fs_settings[2]
		self.fstop = fs_settings[3]
		self.fstep = fs_settings[4]
		self.i_adapt = bool(fs_settings[5])

		# Advanced
		self.errmax = fs_settings[6]
		self.errinc = fs_settings[7]
		self.ninc = int(fs_settings[8])
		self.ioutall = bool(fs_settings[9])

class ResultsExport:
	"""
		Class to deal with handling of the results that are created
	"""
	def __init__(self, pth, name):
		"""
			Initialises, if path passed then results saved there
		:param str pth:  Path to where the results should be saved
		:param str name:  Name that should be used for the results file
		"""
		self.logger = constants.logger

		self.parent_pth = pth
		self.results_name = name

		self.results_folder = str()

	def create_results_folder(self):
		"""
			Routine creates the sub_folder which will contain all of the detailed results
		:return None:
		"""

		self.results_folder = os.path.join(self.parent_pth, self.results_name)

		# Check if already exists and if not create it
		if not os.path.exists(self.results_folder):
			os.mkdir(self.results_folder)

		return None


