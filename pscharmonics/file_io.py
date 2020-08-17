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
import pscharmonics.constants as constants
import glob
import pandas as pd
import numpy as np
import scipy.spatial
import collections
import math
import shutil
import time
import xlsxwriter
import xlsxwriter.utility

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

	if df.empty:
		# If DataFrame is empty then just return empty DataFrame and with updated flag = False
		df_updated = df
		updated = False
	else:
		# Group data frame by key value

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
	:param logger.Logger, logger:  Handle to the logger
	:param tuple thresholds: (int, int) (warning, delete) where
								warning is the number at which a warning message is displayed
								delete is where deleting of the files starts
	:return int num_deleted:
	"""
	threshold_warning = thresholds[0]
	threshold_delete = thresholds[1]

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
		if threshold_warning < num_files < threshold_delete:
			num_deleted = 0
			logger.warning(
				(
					'There are {} files in the folder {}, when this number gets to {} the oldest will be deleted'
				).format(num_files, pth, threshold_delete)
			)
		# Check if more than delete level
		elif num_files >= threshold_delete:
			logger.warning(
				(
					'There are {} files in the folder {}, the oldest will be reduced to get this number down to {}'
				).format(num_files, pth, threshold_warning)
			)
			num_to_delete = num_files - threshold_warning

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
					'{} files in folder {} is less than the warning threshold {} and delete threshold {}'
				).format(num_files, pth, threshold_warning, threshold_delete)
			)

	return num_deleted

class ExtractResults:
	"""
		Values defined during import
	"""
	include_convex = False  # type: bool

	def __init__(self, target_file, search_paths):
		"""
			Process the extraction of the results
		:param str target_file:  Target file to save results to
		:param tuple search_paths:  List of folder to search for the results files to export
		"""
		self.logger = constants.logger

		df, extract_vars = self.combine_multiple_runs(search_paths=search_paths)

		# Function will calculate the convex hull for the R and X values at each node in this DataFrame.
		# Initially False but set to True during importing of multiple runs if appropriate
		df_convex = pd.DataFrame()
		if self.include_convex:
			if constants.PowerFactory.pf_r1 and extract_vars and constants.PowerFactory.pf_x1 in extract_vars:
				# TODO: Target frequencies should be provided as an input, still to be completed
				loci_settings = LociSettings()
				target_freq = loci_settings.freq_bands
				percentage_to_exclude = loci_settings.exclude
				df_convex = calculate_convex_vertices(
					df=df, frequency_bounds=target_freq, percentage_to_exclude=percentage_to_exclude
				)
			else:
				self.logger.warning(
					(
						'Not able to produce ConvexHull because self impedance R ({}) and X ({}) values were not '
						'collected during the study.  Only the following values were collected:\n\t{}'
					).format(constants.PowerFactory.pf_r1, constants.PowerFactory.pf_x1, '\n\t-'.join(extract_vars))
				)

		self.extract_results(pth_file=target_file, df=df, vars_to_export=extract_vars, df_convex=df_convex)

	# noinspection PyMethodMayBeStatic
	def combine_multiple_runs(self, search_paths, drop_duplicates=True):
		"""
			Function will combine multiple results extracts into a single results file
		:param tuple search_paths:  List of folders which contain the results files to be combined / extracted
									each folder must contain raw .csv results exports + a inputs
		:param bool drop_duplicates:  (Optional=True) - If set to False then duplicated columns will be included in the output
		:return pd.DataFrame df, list vars_to_export:
					Combined results into single dataframe,
					list of variables for export
		"""
		c = constants.Results
		logger = constants.logger

		logger.info(
			'Importing all results files in following list of folders: \n\t{}'.format('\n\t'.join(search_paths))
		)

		# Loop through each folder, obtain the files and produce the DataFrames

		all_dfs = []
		vars_to_export = []

		# Loop through each folder, import the inputs sheet and results files
		for folder in search_paths:
			# Import results into a single dataframe
			combined = PreviousResultsExport(pth=folder)
			all_dfs.append(combined.df)

			# Include list of variables for export
			vars_to_export.extend(combined.inputs.settings.get_vars_to_export())

			# Determine whether include_convex is set to True or False and update overall setting accordingly
			# will latch to True if any of the imported files have include_convex and then any errors during processing
			# are dealt with accordingly
			self.include_convex = self.include_convex or combined.inputs.settings.include_convex


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
				logger.warning(('The input data ets had duplicated columns and '
								'therefore some have been removed.\n'
								'{} columns have been removed')
							   .format(original_shape[1]-new_shape[1]))

			else:
				logger.debug('No duplicated data in results files imported from: {}'
							 .format(search_paths))
			if new_shape[0] != original_shape[0]:
				raise SyntaxError('There has been an error in the processing and some rows have been deleted.'
								  'Check the script')
		else:
			logger.debug('No check for duplicates carried out')

		# Sort the results so that in study_case name order
		df.sort_index(axis=1, level=[c.lbl_Reference_Terminal, c.lbl_Terminal, c.lbl_StudyCase], inplace=True)

		return df, vars_to_export

	def extract_results(self, pth_file, df, vars_to_export, df_convex, plot_graphs=True):
		"""
			Extract results into workbook with each result on separate worksheet
		:param str pth_file:  File to save workbook to
		:param pd.DataFrame df:  Pandas dataframe to be extracted
		:param list vars_to_export:  List of variables to export based on Inputs class
		:param pd.DataFrame df_convex:  Pandas DataFrame with the boundaries of the ConvexHull data points
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
		self.logger.info('\tExporting results for {} nodes'.format(num_nodes))

		# Will only include index and header labels if True
		# include_index = col <= c.start_col
		include_index = True

		# Export to excel with a new sheet for each node
		i=0
		try:
			with pd.ExcelWriter(pth_file, engine='xlsxwriter') as writer:
				for node_name, _df in list_dfs:
					self.logger.info('\t - \t {}/{} Exporting node {}'.format(i+1, num_nodes, node_name))
					i += 1
					col = c.start_col

					for var in vars_to_export:
						# Extract DataFrame with just these values
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

					# Once all main results have been exported ConvexHull points for each node are added
					if not df_convex.empty:
						# Based on the raw dataset, start row and start column determine the Excel row and columns
						# that cover the raw R and X values for the calculated impedance loci
						# TODO: Target frequencies should be provided as an input, still to be completed
						target_freq = LociSettings().freq_bands
						#
						raw_x_data, raw_y_data = get_raw_data_excel_references(
							sht_name=node_name,
							df=_df, start_row=start_row, start_col=c.start_col,
							target_frequencies=target_freq)

						# Determine which row to start the convex hull on, taking into consideration the number of
						# rows occupied by the DataFrame
						row_convex = start_row + len(_df) + _df.columns.nlevels + c.row_spacing + 1
						# Get ConvexValues for this node in particular
						df_node = df_convex.loc[:, df_convex.columns.get_level_values(level=c.lbl_Reference_Terminal)==node_name]

						# Results are exported
						df_node.to_excel(writer, merge_cells=True,
										 sheet_name=node_name,
										 startrow=row_convex, startcol=c.start_col,
										 header=include_index, index_label=False)

						# TODO: Write routine to include a graph showing the convex hull data points
						# Add loci plots
						self.add_loci_graphs(
							writer=writer,
							sheet_name=node_name,
							plot_names=df_node.columns.get_level_values(level=c.lbl_Harmonic_Order),
							row_labels=row_convex + 1,
							row_start=row_convex + df_node.columns.nlevels + 1,
							num_rows=df_node.shape[0],
							col_start=c.start_col+1,
							num_cols=df_node.shape[1],
							raw_x_data=raw_x_data,
							raw_y_data=raw_y_data
						)


		except PermissionError:
			self.logger.critical(
				(
					'Unable to write to excel workbook {} since it is either already open or you do not have the '
					'appropriate permissions to save here.  Please check and then rerun.'
				).format(pth_file)
			)
			raise PermissionError('Unable to write to workbook')

		self.logger.info('Completed exporting of results to {}'.format(pth_file))

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

	def add_loci_graphs(self, writer, sheet_name, plot_names, row_labels, row_start, num_rows,
				  col_start, num_cols, raw_x_data, raw_y_data):
		"""
			Add graphs for ConvexHull to spreadsheet

			Two plots are produced for each dataset except the overall plot which includes all Loci:
				Plot 1 = Loci curve + raw data
				Plot 2 = Loci curve alone

		:param pd.ExcelWriter writer:  Handle for the workbook that will be controlling the excel instance
		:param str sheet_name: Name of sheet to add graph to
		:param list plot_names: List of names for the plots
		:param int row_labels: Row which contains series labels
		:param int row_start: Start row for results to plot
		:param int num_rows: Number of rows to include
		:param int col_start:  Column number first set of data is included in
		:param int num_cols:  Number of columns all data is contained within
		:param dict raw_x_data:  Dictionary of excel function for raw R data of points for each plot
		:param dict raw_y_data:  Dictionary of excel function for raw X data of points for each plot
		:return:
		"""
		self.logger.debug('Adding impedance loci plots for node: {}'.format(sheet_name))

		c = constants.Results
		color_map = c().get_color_map()

		# Get handles
		wkbk = writer.book
		sht = writer.sheets[sheet_name]

		# Calculate the row number for the end of the dataset
		max_row = row_start+num_rows

		# Empty lists are populated with all of the charts so that they can neatly be added to the
		# excel plots
		charts = list()
		charts_raw = list()

		# Default colour to use for LOCI plots
		color_i = 1

		# Create a single chart which contains all values and then an individual chart for each harmonic order
		chrt_master = wkbk.add_chart(c.chart_type)
		chrt_master.set_title({'name':'{}'.format(sheet_name),
							   'name_font':{'size':c.font_size_chart_title}})

		# The overall chart does not have the raw data added as it would be too cluttered
		charts.append(chrt_master)

		# Remove every other element in plot_names since the plot titles are duplicated
		plot_names = plot_names[::2]

		# Loop through each pair of R / X values and add to excel plot
		for i, col_r in enumerate(range(col_start, col_start+num_cols, 2)):
			# Iterate color so a new plot is shown
			color_i += 1

			# X values are 1 column over from R values
			col_x = col_r + 1

			# Get chart name from row with description of harmonic order
			chart_name = plot_names[i]
			self.logger.debug('\t Creating loci plot for {}-{}'.format(sheet_name, chart_name))

			# Add a new chart to the workbook to contain the overall loci plot
			chrt_individual = wkbk.add_chart(c.chart_type)
			chrt_individual.set_title({'name':'{} - {}'.format(sheet_name, chart_name),
							'name_font':{'size':c.font_size_chart_title}})

			# Add a new chart to the workbook to also contain the raw points as well as the loci curve
			chrt_raw = wkbk.add_chart(c.chart_type)  # type: pd.ExcelWriter.Chart
			chrt_raw.set_title(
				{'name':'{} - {} (including raw points)'.format(sheet_name, chart_name),
				 'name_font':{'size':c.font_size_chart_title}}
			)

			# Add charts to list for future processing
			charts.append(chrt_individual)
			charts.append(chrt_raw)

			for chrt in (chrt_individual, chrt_master, chrt_raw):

				# Add to plot
				chrt.add_series({
					'name': [sheet_name, row_labels, col_r],
					'categories': [sheet_name, row_start, col_r, max_row, col_r],
					'values': [sheet_name, row_start, col_x, max_row, col_x],
					'marker': {'type': 'none'},
					'line':  {'color': color_map[color_i],
					 		  'width': c.line_width}
				})

			# Add raw data points as a new series to the chart if not empty dictionaries
			if raw_x_data and raw_y_data:
				chrt_raw.add_series({
					'name': 'Raw Points',
					'categories': raw_x_data[chart_name],
					'values': raw_y_data[chart_name],
					'marker': {'type': c.marker_type, 'size': c.market_size},
					'line': {'none': True}
				})



		col_number = 0
		row_number = 0
		for i, chrt in enumerate(charts):
			# Counter to grid the loci plots so they have 4 vertical in each column
			if row_number > c.loci_plots_vertically-1:
				row_number = 0
				col_number += 1
			elif i > 0:
				row_number += 1

			# Add axis labels
			chrt.set_x_axis({'name': c.lbl_Resistance, 'label_position': c.lbl_position,
							 'major_gridlines': c.grid_lines})
			chrt.set_y_axis({'name': c.lbl_Reactance, 'label_position': c.lbl_position,
							 'major_gridlines': c.grid_lines})

			# Set chart size
			chrt.set_size({'width': c.chrt_loci_width, 'height': c.chrt_loci_height})

			# Add the legend to the chart
			sht.insert_chart(row_start+(c.chrt_vert_space*row_number), c.chrt_col+(c.loci_chrt_space*col_number), chrt)

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

	# noinspection PyMethodMayBeStatic
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

			# If greater than 1 then split into equal size groups with base case plot at starting point
			if number_of_plots > 1:
				steps = math.ceil(len(list_of_cols) / number_of_plots)
				a = [list_of_cols[i * steps:(i + 1) * steps] for i in range(int(math.ceil(len(list_of_cols) / steps)))]

				for i, group in enumerate(a):
					graphs['{}({})'.format(title, i + 1)] = group
			else:
				graphs['{}'.format(title)] = list_of_cols

		return graphs

class PreviousResultsExport:
	""" Used for importing the settings and previously exported results """
	def __init__(self, pth):
		"""

		:param str pth:  path that will contain the input files
		"""

		self.logger = constants.logger

		self.search_pth = pth
		self.logger.debug('Processing results saved in: {}'.format(self.search_pth))

		# Constant declarations
		self.study_type = str()

		# Get inputs used
		self.inputs = self.get_input_values()

		# Get a single DataFrame for all results
		self.df = self.import_all_results(study_type=constants.Results.study_fs)
		self.logger.debug('All results for folder: {} imported'.format(self.search_pth))

	def get_input_values(self):
		"""
			Function to import the inputs file found in a folder (only a single file is expected to be found)
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
		if cont_name:
			full_name = '{}_{}'.format(sc_name, cont_name)
		else:
			full_name = sc_name

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

		# Combine data frames into one and return
		df = pd.concat([df, df_mutual], axis=1)

		# Obtain nominal voltage for each terminal
		idx = pd.IndexSlice
		# Add in row to contain the nominal voltage
		df.loc[c.idx_nom_voltage] = str()

		dict_nom_voltage = dict()

		# Find the nominal voltage for each terminal (if exists)
		for term, df_sub in df.groupby(axis=1, level=c.lbl_Reference_Terminal):
			idx_filter = idx[:,:,:,:,:,constants.PowerFactory.pf_nom_voltage]
			try:
				# Obtain nominal voltage and then set row values appropriately to include in results
				nom_voltage = df_sub.loc[:,idx_filter].iloc[0,0]
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

		# Check if intact contingency being applied
		if constants.Contingencies.intact in file_name:
			cont_name = constants.Contingencies.intact
		else:
			# Find which contingency is considered
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
				if (
						'{}.{}'.format(term.substation, c.pf_substation) == var_sub and
						'{}.{}'.format(term.terminal, c.pf_terminal) == var_term
				):
					ref_terminal = term.name
					break

			if ref_terminal == '':
				self.logger.critical(
					(
						'Substation {} in the results does not appear in the inputs list of substations with the '
						'associated terminal {}.  List of inputs are:\n\t{}'
					).format(var_sub, var_term, '\n\t'.join([term.name for term in self.inputs.terminals.values()]))
				)
				raise ValueError(
					(
						'Substation {} not found in inputs {} but has appeared in results'
					).format(var_sub, self.search_pth)
				)

		return var_name, ref_terminal

	# noinspection PyMethodMayBeStatic
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
		# Option to decide whether to include ConvexHull in excel spreadsheets, default value is False
		self.include_convex = False

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
			self.results_name = self.process_results_name()
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

		self.include_convex = self.process_booleans(key=self.c.include_convex)

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

	def process_results_name(self, def_value=constants.StudySettings.def_results_name):
		"""
			Process the results file name that has been provided and if a blank value is provided then use the default
			value
		:param str def_value:  Default file name to use
		:return str res_name:
		"""
		# Get folder from DataFrame, if empty or invalid path then use default folder
		res_name = str(self.df.loc[self.c.results_name])

		# Check if any results file has been provided and then append the correct extension and UID
		# value
		if not res_name:
			# If no folder provided then use default value
			res_name = '{}_{}.xlsx'.format(def_value, constants.uid)
		else:
			res_name = '{}{}.xlsx'.format(res_name, constants.uid)

		# Check if target file already exists and warn user that results will be overwritten
		overall_res_pth = os.path.join(self.export_folder, res_name)
		if os.path.isfile(overall_res_pth):
			# Number of seconds to delay before deleting file
			delay = 5.0
			self.logger.warning(
				'The target results file <{}> already exists and will be deleted in {} seconds'.format(overall_res_pth, delay)
			)
			time.sleep(delay)
			os.remove(overall_res_pth)

		return res_name

	def add_folder(self, pth_results_file):
		"""
			Folder creates a results file
		:param str pth_results_file:
		:return:
		"""
		delay_counter = 5

		# Get the full path to the results file
		pth, file_name = os.path.split(pth_results_file)
		folder_name, _ = os.path.splitext(file_name)

		target_folder = os.path.join(pth, folder_name)

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
		try:
			value = self.df.loc[key]
		except KeyError:
			# If key doesn't exist then use False and let user know they are using an invalid Inputs sheet
			value = False
			self.logger.warning(
				(
					'The inputs spreadsheet you are using does not have an input value for {} on the {} worksheet.  This'
					'means it is either an old Inputs worksheet or has been edited.  A value of <{}> has been assumed '
					'instead'
				).format(key, self.sht, value)
			)

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

class LociSettings:
	"""
		TODO: Add in worksheet for Loci settings in terms of frequency bandings around nominal
	"""
	def __init__(self):
		nom_freq = constants.General.nominal_frequency

		# TODO: To be replaced with an input from the excel spreadsheet
		freq_step = 25.0

		# Produces a dictionary to look up frequency bandings for each harmonic number
		self.freq_bands = dict()
		for h in range(2, 51):
			start_freq = h * nom_freq - freq_step
			stop_freq = h * nom_freq + freq_step
			self.freq_bands[h] = (start_freq, stop_freq)

		# TODO: To be replaced with an input from the excel spreadsheet
		# Percentage of points to exclude
		self.exclude = 0.1

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
			# TODO: Need to combine these together so a single set of results are produced, also need to ensure names
			# TODO: are unique across both sets
			contingency_cmd_breaker, contingencies_breakers = self.process_contingencies(wkbk=wkbk)  # type: str, dict
			contingency_cmd_lines, contingencies_lines = self.process_contingencies(wkbk=wkbk, line_data=True)  # type: str, dict
			self.terminals = self.process_terminals(wkbk=wkbk)
			self.lf_settings = self.process_lf_settings(wkbk=wkbk)
			self.fs_settings = self.process_fs_settings(wkbk=wkbk)

			# Loci setting (TODO: Specific workbook yet to be added)
			self.loci_settings = LociSettings()

		# Combine contingencies from breakers and lines into a single dictionary of contingencies that will be used if
		# a fault case hasn't been defined already
		if contingency_cmd_breaker and contingency_cmd_lines:
			self.logger.warning(
				(
					'You have provided a contingencies command for both breaker outages and line outages.  However, it'
					'is assumed that only a single command would have been defined.\nThese commands were input on the '
					'inputs workbook {} for the worksheets as follows:\n\t{}.\n'
					'The command {} will be used for this study'
				).format(self.pth, contingency_cmd_breaker, contingency_cmd_lines, contingency_cmd_breaker)
			)
			self.contingency_cmd = contingency_cmd_breaker
		elif contingency_cmd_breaker:
			self.logger.debug('Contingencies command based on circuit breaker detailed command {}'.format(contingency_cmd_breaker))
			self.contingency_cmd = contingency_cmd_breaker
		elif contingency_cmd_lines:
			self.logger.debug('Contingencies command based on line outage detailed command {}'.format(contingency_cmd_lines))
			self.contingency_cmd = contingency_cmd_lines
		else:
			self.logger.debug('No command provided for contingency based outages')
			self.contingency_cmd = None


		# Contingencies could relate to either circuit breakers or lines.  The dictionary is populated for both
		self.contingencies = {k: {constants.Contingencies.cb: v} for k,v in contingencies_breakers.items()}

		# Loop through each of the line contingencies and if the key already exists in self.contingencies then add to
		# existing contingency otherwise create a new one.
		existing_contingencies = self.contingencies.keys()
		for key, line_outage in contingencies_lines.items():
			# Check if has already been defined
			if key in existing_contingencies:
				self.logger.debug(
					'The contingency {} is defined with both circuit breaker operations and line outages'.format(key)
				)
				self.contingencies[key][constants.Contingencies.lines] = line_outage
			else:
				# No contingency already exists to just define it with just line outage details
				self.contingencies[key] = {constants.Contingencies.lines: line_outage}

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
		destination = os.path.join(
			self.settings.export_folder, '{}_{}{}'.format(
				constants.StudyInputs.file_name,
				constants.uid,
				file_name
			)
		)

		# Copy to destination
		self.logger.debug('Saving inputs file {} in to export folder: {}'.format(src, destination))
		shutil.copyfile(src=src, dst=destination)

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

	def process_contingencies(self, sht=None, wkbk=None, pth_file=None, line_data=False) -> (str, dict):
		"""
			Function imports the DataFrame of contingencies and then each into its own contingency under a single
			dictionary key.  Each dictionary item contains the relevant outages to be taken in the form
			ContingencyDetails.

			Also returns a string detailing the name of a Contingency file if one is provided

		:param str sht:  (optional) Name of worksheet to use
		:param pd.ExcelFile wkbk:  (optional) Handle to workbook
		:param str pth_file: (optional) File path to workbook
		:param bool line_data: (optional) Set to True if processing line data rather than circuit breaker data
		:return (str, dict) (contingency_cmd, contingencies):  Returns both:
																The command for all contingencies
																A dictionary with each of the outages
		"""
		if line_data:
			data_type = 'Line Outages'
			# If not sheet name has been provided then determine which sheet needs importing
			if sht is None:
				sht = constants.StudyInputs.cont_lines
		else:
			data_type = 'Circuit Breaker Operations'
			# If not sheet name has been provided then determine which sheet needs importing
			if sht is None:
				sht = constants.StudyInputs.cont_breakers

		self.logger.debug('Importing contingencies data for {} type'.format(data_type))

		# Import workbook as dataframe
		if wkbk is None:
			wkbk = self.load_workbook(pth_file=pth_file)

		# Confirm sheet exists in workbook and if not raise a warning and return
		if sht not in wkbk.sheet_names:
			self.logger.error(
				(
					'The worksheet named {} cannot be found in the provided inputs workbook {} and therefore the '
					'contingencies specified by type {} cannot be processed.  You may be using an old workbook in '
					'which case this will not be an issue'
				).format(sht, self.pth, data_type)
			)
			# Return an empty string and empty dictionary
			return str(), dict()
		else:
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
				# Process contingencies taking into consideration whether processing circuit breakers or lines
				cont = ContingencyDetails(list_of_parameters=item.values, line_data=line_data)
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
	def __init__(self, list_of_parameters, line_data=False):
		"""
			Single row of contingency parameters imported from spreadsheet are defined into a class for easy lookup
			This relates to contingencies which are identified by their breakers rather than names of lines
		:param list list_of_parameters:  Single row of inputs
		:param bool line_data:  If being provided with inputs that relate to line names rather than circuit breakers
		"""
		# Status flag set to True if cannot be found, contingency fails, etc.
		self.not_included = False

		self.name = list_of_parameters[0]

		# Determines whether Contingency relates to lines or circuit breakers
		self.line_data = line_data

		# Initialise empty lists to be populated with details of the outage circuits or lines
		self.couplers = list()
		self.lines = list()

		if self.line_data:
			# Inputs are names of lines to be outage and so populate lines with details
			for line, status in zip(*[iter(list_of_parameters[1:])]*2):
				if line != '' and not pd.isna(line):
					new_line = LineDetails(str(line), status)
					self.lines.append(new_line)

		else:
			# Inputs relate to circuit breaker identification
			for substation, breaker, status in zip(*[iter(list_of_parameters[1:])]*3):
				if substation != '' and breaker != '' and str(breaker) != 'nan':
					new_coupler = CouplerDetails(substation, breaker, status)
					self.couplers.append(new_coupler)

		# If contingency has been defined then needs to be included in results
		# # Check if this contingency relates to the intact system in which case it will be skipped
		if self.name == constants.StudyInputs.base_case or (len(self.couplers) == 0 and len(self.lines)==0):
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
		# Check if substation already has substation type ending and if it needs removing since it is added as part of
		# search algorithm
		if str(substation).endswith(constants.PowerFactory.pf_substation):
			substation = substation.replace('.{}'.format(constants.PowerFactory.pf_substation), '')

		# Check if substation already has substation type ending and if it needs removing since it is added as part of
		# search algorithm
		if str(breaker).endswith(constants.PowerFactory.pf_coupler):
			breaker = breaker.replace('.{}'.format(constants.PowerFactory.pf_coupler),'')

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

class LineDetails:
	def __init__(self, line, status):
		"""
			Define the line and status for outages relating to lines
		:param str line:  Name of line
		:param str status: Status of line (in service or out of service)
		"""
		logger = constants.logger
		# Check if line already has type ending and if it needs removing since it is added as part of search algorithm
		if line.endswith(constants.PowerFactory.pf_line):
			line = line.replace('.{}'.format(constants.PowerFactory.pf_line), '')
		elif line.endswith(constants.PowerFactory.pf_branch):
			line = line.replace('.{}'.format(constants.PowerFactory.pf_branch), '')

		# Confirm that status for operation is either true or false
		if status == constants.StudyInputs.in_service:
			status = 1
		elif status == constants.StudyInputs.out_of_service:
			status = 0
		else:
			# Post a warning message for an unexpected input
			logger.warning(
				(
					'The state requested for the line <{}> has a value of {} which is not an expected value of '
					'{} or {}.  The operation is assumed to be {} for this study but the user should check '
					'that is what they intended'
				).format(
					line, status,
					constants.StudyInputs.in_service, constants.StudyInputs.out_of_service,
					constants.StudyInputs.out_of_service
				)
			)
			status = 0

		self.line = line
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
		# Check if substation already has substation type ending and if it needs removing since it is added as part of
		# search algorithm
		if str(substation).endswith(c.pf_substation):
			substation = substation.replace('.{}'.format(c.pf_substation), '')

		if str(terminal).endswith(c.pf_terminal):
			terminal = terminal.replace('.{}'.format(c.pf_terminal), '')

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
		self.iopt_net = int(load_flow_settings[0])  # Calculation method (0 Balanced AC, 1 Unbalanced AC, DC)
		self.iopt_at = int(load_flow_settings[1])  # Automatic Tap Adjustment
		self.iopt_asht = int(load_flow_settings[2])  # Automatic Shunt Adjustment

		# Added in Automatic Tapping of PSTs with default values
		self.iPST_at = int(load_flow_settings[3])  # Automatic Tap Adjustment of Phase Shifters

		self.iopt_lim = int(load_flow_settings[4])  # Consider Reactive Power Limits
		self.iopt_limScale = int(load_flow_settings[5])  # Consider Reactive Power Limits Scaling Factor
		self.iopt_tem = int(load_flow_settings[6])  # Temperature Dependency: Line Cable Resistances (0 ...at 20C, 1 at Maximum Operational Temperature)
		self.iopt_pq = int(load_flow_settings[7])  # Consider Voltage Dependency of Loads
		self.iopt_fls = int(load_flow_settings[8])  # Feeder Load Scaling
		self.iopt_sim = int(load_flow_settings[9])  # Consider Coincidence of Low-Voltage Loads
		self.scPnight = float(load_flow_settings[10])  # Scaling Factor for Night Storage Heaters

		# Active Power Control
		self.iopt_apdist = int(load_flow_settings[11])  # Active Power Control (0 as Dispatched, 1 According to Secondary Control,
		# 2 According to Primary Control, 3 According Inertia)
		self.iopt_plim = int(load_flow_settings[12])  # Consider Active Power Limits
		self.iPbalancing = int(load_flow_settings[13])  # (0 Ref Machine, 1 Load, Static Gen, Dist slack by loads, Dist slack by Sync,

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


		self.phiini = float(load_flow_settings[15])  # Angle

		# Advanced Options
		self.i_power = int(load_flow_settings[16])  # Load Flow Method ( NR Current, 1 NR (Power Eqn Classic)
		self.iopt_notopo = int(load_flow_settings[17])  # No Topology Rebuild
		self.iopt_noinit = int(load_flow_settings[18])  # No initialisation
		self.utr_init = int(load_flow_settings[19])  # Consideration of transformer winding ratio
		self.maxPhaseShift = float(load_flow_settings[20])  # Max Transformer Phase Shift
		self.itapopt = int(load_flow_settings[21])  # Tap Adjustment ( 0 Direct, 1 Step)
		self.krelax = float(load_flow_settings[22])  # Min Controller Relaxation Factor

		self.iopt_stamode = int(load_flow_settings[23])  # Station Controller (0 Standard, 1 Gen HV, 2 Gen LV
		self.iopt_igntow = int(load_flow_settings[24])  # Modelling Method of Towers (0 With In/ Output signals, 1 ignore couplings, 2 equation in lines)
		self.initOPF = int(load_flow_settings[25])  # Use this load flow for initialisation of OPF
		self.zoneScale = int(load_flow_settings[26])  # Zone Scaling ( 0 Consider all loads, 1 Consider adjustable loads only)

		# Iteration Control
		self.itrlx = int(load_flow_settings[27])  # Max No Iterations for Newton-Raphson Iteration
		self.ictrlx = int(load_flow_settings[28])  # Max No Iterations for Outer Loop
		self.nsteps = int(load_flow_settings[29])  # Max No Iterations for Number of steps

		self.errlf = float(load_flow_settings[30])  # Max Acceptable Load Flow Error for Nodes
		self.erreq = float(load_flow_settings[31])  # Max Acceptable Load Flow Error for Model Equations
		self.iStepAdapt = int(load_flow_settings[32])  # Iteration Step Size ( 0 Automatic, 1 Fixed Relaxation)
		self.relax = float(load_flow_settings[33])  # If Fixed Relaxation factor
		self.iopt_lev = int(load_flow_settings[34])  # Automatic Model Adaptation for Convergence

		# Outputs
		self.iShowOutLoopMsg = int(load_flow_settings[35])  # Show 'outer Loop' Messages
		self.iopt_show = int(load_flow_settings[36])  # Show Convergence Progress Report
		self.num_conv = int(load_flow_settings[37])  # Number of reported buses/models per iteration
		self.iopt_check = int(load_flow_settings[38])  # Show verification report
		self.loadmax = float(load_flow_settings[39])  # Max Loading of Edge Element
		self.vlmin = float(load_flow_settings[40])  # Lower Limit of Allowed Voltage
		self.vlmax = float(load_flow_settings[41])  # Upper Limit of Allowed Voltage
		self.iopt_chctr = int(load_flow_settings[42])  # Check Control Conditions

		# Load Generation Scaling
		self.scLoadFac = float(load_flow_settings[43])  # Load Scaling Factor
		self.scGenFac = float(load_flow_settings[44])  # Generation Scaling Factor
		self.scMotFac = float(load_flow_settings[45])  # Motor Scaling Factor

		# Low Voltage Analysis
		self.Sfix = float(load_flow_settings[46])  # Fixed Load kVA
		self.cosfix = float(load_flow_settings[47])  # Power Factor of Fixed Load
		self.Svar = float(load_flow_settings[48])  # Max Power Per Customer kVA
		self.cosvar = float(load_flow_settings[49])  # Power Factor of Variable Part
		self.ginf = float(load_flow_settings[50])  # Coincidence Factor
		self.i_volt = int(load_flow_settings[51])  # Voltage Drop Analysis (0 Stochastic Evaluation, 1 Maximum Current Estimation)

		# Advanced Simulation Options
		self.iopt_prot = int(load_flow_settings[52])  # Consider Protection Devices ( 0 None, 1 all, 2 Main, 3 Backup)
		self.ign_comp = int(load_flow_settings[53])  # Ignore Composite Elements

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


def find_convex_vertices(x_values, y_values):
	"""
		Finds the ConvexHull that bounds around the provided x and y values
	:param tuple x_values: X axis values to be considered
	:param tuple y_values: Y axis values to be considered
	:return tuple corners: (x / y points for each corner
	"""
	c = constants.PowerFactory

	# Filter out values which are outside of allowed range
	x_points = list()
	y_points = list()
	for x,y in zip(x_values, y_values):
		if 0 < abs(x) < c.max_impedance and 0 < abs(y) < c.max_impedance:
			x_points.append(x)
			y_points.append(y)

	# Convert provided data points into ndarray
	points = np.column_stack((x_points, y_points))

	# Confirm the arrays not empty and must be greater than 2 to actually calculate some points
	if points.shape[0] > 2:
		# Determines the ConvexHull and extracts the vertices
		hull = scipy.spatial.ConvexHull(points=points)
		corners = hull.vertices

		# Convert the vertices to represent the x and y values and include the origin
		x_corner = list(points[corners, 0])
		x_corner.append(x_corner[0])
		y_corner = list(points[corners, 1])
		y_corner.append(y_corner[0])
	elif points.shape[0] > 0:
		# If length of 1 or 2 then just return those points
		x_corner = x_points
		y_corner = y_points
	else:
		# If no points then return empty list
		x_corner = list()
		y_corner = list()

	# Return tuple of R and X values
	return x_corner, y_corner

def calculate_convex_vertices(df, frequency_bounds, percentage_to_exclude, nom_frequency=50.0):
	"""
		Will loop through the provided DataFrame and calculate the convex hull that bounds the R and X
		values for each.

		The minimum and maximum values provided for each frequency bound are included

		A specific percentage of points will be excluded from the impedance loci (i.e. the top 10% of values)
	:param pd.DataFrame df:  This is the DataFrame containing R and X values that will then be processed for extraction
	:param dict frequency_bounds:  For each harmonic number provides the starting and stopping frequency in the format {
									str harmonic number: (float minimum frequency, float maximum frequency)
									}
	:param float percentage_to_exclude:  Percentage of maximum points to exclude from the dataset
	:param float nom_frequency:  Nominal frequency = 50.0 Hz
	:return pd.DataFrame df_convex:  Returns a DataFrame in the same arrangement as the supplied DataFrame but with the
									corners for each vertices
	"""
	# Confirm that the percentage to exclude value has been provided as a float rather than a percentage
	if percentage_to_exclude > 1.0:
		constants.logger.warning(
			(
				'The percentage to exclude value ({:.1f}%) is greater than 100 % and therefore has been input as a '
				'percentage rather than a float.  To resolve this a value of {:.1f} % has been considered instead.'
			).format(percentage_to_exclude*100.0, percentage_to_exclude)
		)
		# Update value to divide by 100.0
		percentage_to_exclude = percentage_to_exclude / 100.0

	# Obtain constants
	c = constants.Results

	# Populated with the Convex Hull points for each node
	dict_convex = dict()

	# Loop through each node
	for node_name, df_node in df.groupby(level=c.lbl_Reference_Terminal, axis=1):
		# Obtain df_z so can extract the largest numbers and exclude them from the filtering
		df_z = df_node.loc[:, df_node.columns.get_level_values(level=c.lbl_Result)==constants.PowerFactory.pf_z1]

		# Continue if no data exists for this node or if it is marked for deleting then no need to process any further
		if node_name == constants.Results.lbl_to_delete or df_z.empty:
			continue

		# Extract r values and x values
		df_r = df_node.loc[:, df_node.columns.get_level_values(level=c.lbl_Result)==constants.PowerFactory.pf_r1]
		df_x = df_node.loc[:, df_node.columns.get_level_values(level=c.lbl_Result)==constants.PowerFactory.pf_x1]

		# Empty Dictionary gets populated with the required harmonic numbers
		dict_harms = dict()

		# Loop through each harm_number taking steps based on harm_groups and then use the middle of the range
		for h, freq_range in frequency_bounds.items():
			if max(freq_range) <= nom_frequency:
				# No R / X data extracted for fundamental
				continue
			# Get the extremes of the allowable frequency range
			min_f_range = min(freq_range)
			max_f_range = max(freq_range)
			descriptor = 'h = {}  ({} - {} Hz)'.format(h, min_f_range, max_f_range)
			idx_selection = (df_r.index >= min_f_range) & (df_r.index <= max_f_range)

			# Confirm that there are actually any indexes for this harmonic number and if so extract results
			if any(idx_selection):
				# Extract the Z1 values specific to this frequency range and identify the index values for those which
				# exceeded the allowed percentile
				z_harm = df_z[idx_selection].values.ravel()
				percentile_value = np.percentile(a=z_harm, q=(1-percentage_to_exclude)*100.0)
				# Find index values for all values that are less than the percentile value
				idx_keep = z_harm<=percentile_value

				if not any(idx_keep):
					constants.logger.error(
						(
							'For node {} with harmonic number {} covering the frequencies {:.1f} to {:.1f} Hz and '
							'excluding the top {:.1f} % has resulted in no values being kept.'
						).format(node_name, h, min_f_range, max_f_range, percentage_to_exclude*100.0)
					)
				else:
					# Extract the 2D DataFrame of values into a 1D numpy array and only keep those values which are less
					# then the percentile values
					# https://stackoverflow.com/questions/13730468/from-nd-to-1d-arrays
					r_harm = df_r[idx_selection].values.ravel()[idx_keep]
					x_harm = df_x[idx_selection].values.ravel()[idx_keep]

					vertices = find_convex_vertices(x_values=r_harm, y_values=x_harm)

					# Create a new DataFrame with the vertices as columns
					df_single_harm = pd.DataFrame(data=np.array(vertices).T, columns=(constants.PowerFactory.pf_r1, constants.PowerFactory.pf_x1))
					dict_harms[descriptor] = df_single_harm

		# Combine all into a single DataFrame
		df_all_harms = pd.concat(dict_harms.values(), keys=dict_harms.keys(), axis=1)
		dict_convex[node_name] = df_all_harms

	# Combine DataFrames for each node into a single DataFrame
	df_convex = pd.concat(
		dict_convex.values(), keys=dict_convex.keys(), axis=1, names=(
			constants.Results.lbl_Reference_Terminal,
			constants.Results.lbl_Harmonic_Order,
			constants.Results.lbl_Result
		)
	)

	return df_convex

def get_raw_data_excel_references(sht_name, df, start_row, start_col, target_frequencies):
	"""
		Function returns the cell references which contain all of the raw data points associated with each convex hull
	:param str, sht_name:  Name of the worksheet being written to
	:param pd.DataFrame, df:  DataFrame about to be written to excel
	:param int start_row:  Starting row number that will be plotted
	:param int start_col:  Starting column number that data will begin in
	:param dict target_frequencies:  Dictionary of the frequencies associated with each harmonic number
	:return (dict, dict), (raw_x, raw_y):  Dictionary of the values to be returned
	"""

	number_of_header_rows = df.columns.nlevels
	c = constants.Results

	# Get the number of columns associated with the R data
	number_of_columns = len(df.loc[:, df.columns.get_level_values(level=c.lbl_Result)==constants.PowerFactory.pf_r1].columns)

	# Start column gets adjusted to account for the Z1 impedance data that is always plotted in the first columns
	start_col = start_col + number_of_columns + 2

	# Determine the start and end columns for each dataset
	r_start_col = xlsxwriter.utility.xl_col_to_name(start_col + 1)
	r_end_col = xlsxwriter.utility.xl_col_to_name(start_col + number_of_columns)

	x_start_col = xlsxwriter.utility.xl_col_to_name(start_col + 2 + number_of_columns + 1)
	x_end_col = xlsxwriter.utility.xl_col_to_name(start_col + 2 + number_of_columns + number_of_columns)

	raw_x = dict()
	raw_y = dict()

	# Loop through each harmonic order and find all of the rows within each frequency range
	for h, freq_limits in target_frequencies.items():
		# Get rows that correspond to this frequency range
		idx_selection = (df.index >= min(freq_limits)) & (df.index <= max(freq_limits))
		row_numbers = np.where(idx_selection==True)[0]

		# Convert to include the start and header row numbers (+2 is to account for index header names and extra rows)
		row_numbers = [row+2+number_of_header_rows+start_row for row in row_numbers]

		chart_name = 'h = {}  ({} - {} Hz)'.format(h, min(freq_limits), max(freq_limits))

		# Produce a single string which covers all of the R data, brackets are used to pass as a list
		# of rows into excel so that it handles it as a continuous series
		raw_x[chart_name] = "(" + ",".join([
			"'{}'!${}${}:${}${}".format(
				sht_name, r_start_col, row, r_end_col, row
			) for row in row_numbers
		 ]) + ")"

		raw_y[chart_name] = "(" + ",".join([
			"'{}'!${}${}:${}${}".format(
				sht_name, x_start_col, row, x_end_col, row
			) for row in row_numbers
		]) + ")"

	return raw_x, raw_y
