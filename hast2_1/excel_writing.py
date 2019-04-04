"""
#######################################################################################################################
###											Excel Writing															###
###		Script deals with writing of data to excel and ensuring that a new instance of excel is used for processing	###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###		project JI6973 for EirGrid project PSPF010 - Specialise Support in Power Quality Analysis during 2018		###
###																													###
#######################################################################################################################
"""

import win32com.client              	# Windows COM clients needed for excel etc. if having trouble see notes
import unittest
import re
import os
import time
import math
import numpy as np
import scipy.spatial
import scipy.spatial.qhull
from scipy.spatial import ConvexHull
import itertools
import hast2_1.constants as constants


# Meta Data
__author__ = 'David Mills'
__version__ = '1.4.1'
__email__ = 'david.mills@pscconsulting.com'
__phone__ = '+44 7899 984158'
__status__ = 'In Development - Alpha'

""" Following functions are used purely for processing the inputs of the HAST worksheet """
def add_contingency(row_data):
	"""
		Function to read in the contingency data and save to list
	:param list row_data:
	:return list combined_entry:
	"""
	if len(row_data) > 2:
		aa = row_data[1:]
		station_name = aa[0::3]
		breaker_name = aa[1::3]
		breaker_status = aa[2::3]
		breaker_name1 = ['{}.{}'.format(nam, constants.PowerFactory.pf_coupler) for nam in breaker_name]
		combined_entry = list(zip(station_name, breaker_name1, breaker_status))
		combined_entry.insert(0, row_data[0])
	else:
		combined_entry = [row_data[0], [0]]
	return combined_entry

def add_scenarios(row_data):
	"""
		Function to read in the scenario data and save to list
	:param list row_data:
	:return list combined_entry:
	"""
	combined_entry = [
		row_data[0],
		row_data[1],
		'{}.{}'.format(row_data[2], constants.PowerFactory.pf_case),
		'{}.{}'.format(row_data[3], constants.PowerFactory.pf_scenario)]
	return combined_entry

def add_terminals(row_data):
	"""
		Function to read in the terminals data and save to list
			:param list row_data:
			:return list combined_entry:
			"""
	combined_entry = [
		row_data[0],
		'{}.{}'.format(row_data[1], constants.PowerFactory.pf_substation),
		'{}.{}'.format(row_data[2], constants.PowerFactory.pf_terminal),
		# Third column now contains TRUE or FALSE.  If True then data will be included including
		# transfer impedance from other nodes to this node.  If False then no data will be included.
		row_data[3]]
	return combined_entry

class SubstationFilter:
	"""
		Class for each filter from the HAST import spreadsheet with a new entry for each substation
	"""
	def __init__(self, row_data):
		"""
			Function to read in the filters and save to list
		:param list row_data:  List of values in the form:
			[name to use for filters,
			substation filter belongs to,
			terminal at which filter should be connected,
			type of filter to use (integer based on PF type),
			Q start, Q stop, number of sizes
			freq start, freq stop, number of freq steps,
			quality factor to use,
			parallel resistance (Rp) value to use
			]
		:return list combined_entry:
		"""
		# Variable initialisation
		self.include = True
		self.nom_voltage = 0.0

		# Confirm row data exists
		if row_data[0] is None:
			self.include = False
			return

		# Name to use for filter
		self.name = row_data[0]
		# Substation and terminal within substation that filter should be connected to
		self.sub ='{}.{}'.format(row_data[1], constants.PowerFactory.pf_substation)
		self.term = '{}.{}'.format(row_data[2], constants.PowerFactory.pf_terminal)
		# Type of filter to use
		self.type = constants.PowerFactory.Filter_type[row_data[3]]
		# Q values for filters (start, stop, no. steps)
		self.q_range = list(np.linspace(row_data[4], row_data[5], row_data[6]))
		self.f_range = list(np.linspace(row_data[7], row_data[8], row_data[9]))
		# Quality factor and parallel resistance values to use
		self.quality_factor = row_data[10]
		self.resistance_parallel = row_data[11]

		# Produce lists of each Q step for each frequency so multiple filters can be tested
		self.f_q_values = list(itertools.product(self.f_range, self.q_range))

# REMOVED:  Now considered as a class above
# def add_filters(row_data):
# 	"""
# 		Function to read in the filters and save to list
# 	:param list row_data:
# 	:return list combined_entry:
# 	"""
# 	combined_entry = [
# 		row_data[0],		# Name to use
# 		'{}{}'.format(row_data[1],Constants.pf_substation),		# Substation name + PF identifier
# 		'{}{}'.format(row_data[2],Constants.pf_terminal),		# Terminal to use + PF identifier
# 		Constants.Filter_type[row_data[3]],
# 		[np.linspace(row_data[4], row_data[5], row_data[6])],
# 		[np.linspace(row_data[7], row_data[8], row_data[9])],
# 		row_data[10],
# 		row_data[11]
# 	]
#
# 	return combined_entry

def add_lf_settings(row_data):
	"""
		Function to read in the load flow settings and save to list
	:param list row_data:
	:return list combined_entry:
	"""
	z = row_data
	combined_entry = [
		int(z[0]), int(z[1]), int(z[2]), int(z[3]), int(z[4]), int(z[5]), int(z[6]), int(z[7]),
		int(z[8]),
		float(z[9]), int(z[10]), int(z[11]), int(z[12]), z[13], z[14], int(z[15]), int(z[16]),
		int(z[17]),
		int(z[18]), float(z[19]), int(z[20]), float(z[21]), int(z[22]), int(z[23]), int(z[24]),
		int(z[25]),
		int(z[26]), int(z[27]), int(z[28]), z[29], z[30], int(z[31]), z[32], int(z[33]), int(z[34]),
		int(z[35]), int(z[36]), int(z[37]), z[38], z[39], z[40], z[41], int(z[42]), z[43], z[44], z[45],
		z[46], z[47], z[48], z[49], z[50], z[51], int(z[52]), int(z[53]), int(z[54])]
	return combined_entry

def add_freq_settings(row_data):
	"""
		Function to read in the frequency sweep settings and save to list
	:param list row_data:
	:return list combined_entry:
	"""
	z = row_data
	combined_entry = [z[0], z[1], int(z[2]), z[3], z[4], z[5], int(z[6]), z[7], z[8], z[9],
					  z[10], z[11], z[12], z[13], int(z[14]), int(z[15])]
	return combined_entry

def add_hlf_settings(row_data):
	"""
		Function to read in the harmonic load flow settings and save to list
	:param list row_data:
	:return list combined_entry:
	"""
	z = row_data
	combined_entry = [int(z[0]), int(z[1]), int(z[2]), int(z[3]), z[4], z[5], z[6], z[7],
					  z[8], int(z[9]), int(z[10]), int(z[11]), int(z[12]), int(z[13]), int(z[14])]
	return combined_entry

class Excel:
	""" Class to deal with the writing and reading from excel and therefore allowing unittesting of the
	functions

	Note:  Each call will create a new excel instance
	"""
	def __init__(self, print_info=print, print_error=print):
		"""
			Function initialises a new instance of an excel application
		# TODO: To be replaced with logging handler
		:param builtin_function_or_method print_info:  Handle for print engine used for printing info messages
		:param builtin_function_or_method print_error:  Handle for print engine used for printing error messages
		"""
		# Constants
		# Sheets and starting rows for analysis
		self.analysis_sheets = constants.analysis_sheets
		# IEC limits
		# TODO: Should be moved to a constants or imported from inputs workbook
		# If on input spreadsheet then can be used to test against allocated limits
		self.iec_limits = constants.iec_limits
		self.limits = list(zip(*[self.iec_limits[1]]))

		# Updated with logging handlers once setup finished
		self.log_info = print_info
		self.log_error = print_error

	def __enter__(self):
		# Launch  new excel instance
		self.xl = win32com.client.DispatchEx('Excel.Application')
		# Following code ensures that makepy has been run to obtain the excel constants and defines them
		# TODO: Need to do something to ensure that a new instance is always created so that if excel is opened
		# TODO: whilst that instance is already active it does not get closed.
		_xl = win32com.client.gencache.EnsureDispatch('Excel.Application')
		self.excel_constants = win32com.client.constants
		del _xl
		self.log_info('Excel instance initialised')
		return self

	def __exit__(self, exc_type, exc_value, traceback):
		"""
			Function deals with closing the excel instance once its been created
		:return:
		"""
		# Disable alerts and quit excel
		self.xl.DisplayAlerts = False
		self.xl.Quit()
		self.log_info('excel instance closed')

	def import_excel_harmonic_inputs(self, workbookname):  # Import Excel Harmonic Input Settings
		"""
			Import Excel Harmonic Input Settings
		:param str workbookname: Name of workbook to be imported
		:return analysis_dict: Dictionary of the settings for the analysis work
		"""
		# TODO: Rewrite analysis_dict as class
		# Initialise empty dictionary
		analysis_dict = dict()

		wb = self.xl.Workbooks.Open(workbookname)  # Open workbook
		c = self.excel_constants
		# print(c.xlDown)
		self.xl.Visible = False  # Make excel Visible
		self.xl.DisplayAlerts = False  # Don't Display Alerts

		# Loop through each worksheet defined in <analysis_sheets>
		for x in self.analysis_sheets:
			# Select and activate each worksheet
			ws = wb.Sheets(x[0])  # Set Active Sheet
			# Don't think sheet needs to be activated
			ws.Activate()  # Activate Sheet
			cell_start = x[1]  # Starting Cell

			ws.Range(cell_start).End(c.xlDown).Select()  # Equivalent to shift end down
			row_end = self.xl.Selection.Address
			row_input = []
			current_worksheet = x[0]

			# Code only to be executed for these sheets
			if current_worksheet in constants.PowerFactory.HAST_Input_Scenario_Sheets:
				# if x[0] == "Contingencies" or x[0] == "Base_Scenarios" or x[0] == "Terminals":	# For these sheets
				# Find the starting and ending cells
				cell_start_alph = re.sub('[^a-zA-Z]', '', cell_start)  # Gets the starting cell alpha C5 = C
				cell_start_num = int(re.sub('[^\d.]', '', cell_start))  # Gets the starting cell number C5 = 5
				cell_end = int(re.sub('[^\d.]', '', row_end))  # Gets the ending cell alpha E5 = E
				cell_range_num = list(range(cell_start_num, (cell_end + 1)))  # Gets the ending cell number E5 = 5

				# Check the value below cell is appropriate
				check_value = ws.Range(
					cell_start_alph + str(cell_start_num + 1)).Value  # Check the value below cell called

				if not check_value:
					aces = [cell_start]
				else:
					aces = [cell_start_alph + str(no) for no in cell_range_num]  #

				# Initialise row counter and loop through each row of input data
				count_row = 0
				while count_row < len(aces):
					ws.Range(aces[count_row]).End(c.xlToRight).Select()
					col_end = self.xl.Selection.Address  # Returns address of last cells
					row_value = ws.Range(aces[count_row] + ":" + col_end).Value
					row_value = row_value[0]

					# Routine only for 'Contingencies' worksheet
					if current_worksheet == constants.PowerFactory.sht_Contingencies:
						row_value = add_contingency(row_data=row_value)

					# Routine for Base_Scenarios worksheet
					elif current_worksheet == constants.PowerFactory.sht_Scenarios:
						row_value = add_scenarios(row_data=row_value)

					# Routine for Terminals worksheet
					elif current_worksheet == constants.PowerFactory.sht_Terminals:
						row_value = add_terminals(row_data=row_value)

					# Routine for Filters worksheet
					elif current_worksheet == constants.PowerFactory.sht_Filters:
						row_value = SubstationFilter(row_data=row_value)

						# #row_value = add_filters(row_data=row_value)

					row_input.append(row_value)
					count_row = count_row + 1

			# More efficiently checking which worksheet looking at
			elif current_worksheet in constants.PowerFactory.HAST_Input_Settings_Sheets:
				row_value = ws.Range(cell_start + ":" + row_end).Value
				for item in row_value:
					row_input.append(item[0])
				if current_worksheet == constants.PowerFactory.sht_LF:
					# Process inputs for Loadflow_Settings
					row_input = add_lf_settings(row_data=row_input)

				elif current_worksheet == constants.PowerFactory.sht_Freq:
					# Process inputs for Frequency_Sweep settings
					row_input = add_freq_settings(row_data=row_input)

				elif current_worksheet == constants.PowerFactory.sht_HLF:
					# Process inputs for Harmonic_Loadflow
					row_input = add_hlf_settings(row_data=row_input)

			# Combine imported results in a dictionary relevant to the worksheet that has been imported
			analysis_dict[current_worksheet] = row_input  # Imports range of values into a list of lists

		wb.Close()  # Close Workbook
		return analysis_dict

	def create_workbook(self, workbookname, excel_visible):  # Create Workbook
		"""
			Function creates the workbook for results to be written into
		:param str workbookname: Name to be given to workbook
		:param bool excel_visible:  Constant defining whether excel is visible or not
		:return workbook wb: Handle for workbook that has been created
		"""
		# Create workbook
		self.log_info('Creating workbook {}'.format(workbookname))
		wb = self.xl.Workbooks.Add()

		# Sets excel either visible or invisible depending on constant
		self.xl.Visible = excel_visible  # Make excel Visible
		self.xl.DisplayAlerts = False  # Don't Display Alerts

		# Save workbook
		wb.SaveAs(workbookname)  # Save Workbook
		# Returns handle for workbook and handle for excel application
		return wb

	def create_sheet_plot(self, sheet_name, fs_results, hrm_results, wb,
						  excel_export_rx, excel_export_z, excel_export_hrm,
						  fs_sim, excel_export_z12, excel_convex_hull,
						  hrm_sim):  # Extract information from out file
		"""
			Extract information form powerfactory file and write to workbook
		:param str sheet_name: Name of worksheet
		:param fs_results: Results form frequency scan
		:param hrm_results: Results from harmonic load flow
		:param Excel.Workbook wb: workbook to write data to (_wb used to avoid shadowing)
		:param bool excel_export_rx:  Boolean to determine whether RX data should be exported
		:param bool excel_export_z:  Boolean to determine whether Z data should be exported
		:param bool excel_export_hrm:  Boolean to determine whether harmonic flow data should be exported
		:param bool fs_sim:  Boolean to determine whether frequency scan results should be exported
		:param bool excel_export_z12:  Boolean to determine whether frequency scan results should be exported
		:param bool excel_convex_hull:  Boolean to determine whether the convex hull should be plotted
		:param bool hrm_sim:  Boolean to determine whether harmonic load flow has been carried out
		:return:
		"""
		# Constant declarations
		c = self.excel_constants
		t1 = time.clock()

		# Check if sheet already exists with that name and if it does then find the next suitable name and report change
		# to user
		sheet_name = self.get_sheet_name(sheet_name=sheet_name, wb=wb)

		self.log_info('Creating Sheet: {}'.format(sheet_name))
		# Adds new worksheet
		ws = wb.Worksheets.Add()  # Add worksheet
		ws.Name = sheet_name  # Rename worksheet

		startrow = 2
		startcol = 1
		newcol = 1

		# r_first, r_last, x_first, x_last, z_first, z_last, z_12_first, z_12_last, hrm_endrow, hrm_first, hrm_last = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
		if excel_export_rx:
			startcol = 19
		if excel_export_z or excel_export_hrm:
			startrow = 33
		if excel_export_z and excel_export_hrm:
			startrow = 62

		if fs_sim:
			# Exporting of frequency vs impedance data to excel
			if excel_export_rx or excel_export_z or excel_export_z12:  # Prints the FS Scale
				# End row originally defined based on length of fs_results but since HRLF harmonic results are included
				# up to the 40th harmonic need to take into consideration the max length of those results as well
				if excel_export_hrm:
					endrow = startrow + max(len(fs_results[0]), len(hrm_results[0])) -1
				else:
					endrow = startrow + len(fs_results[0]) - 1
				# Plots the Scale_______________________________________________________________________________________
				scale = fs_results[0]
				scale_end = scale[-1]
				ws.Range(ws.Cells(startrow, startcol), ws.Cells(endrow, startcol)).Value = list(zip(*[fs_results[0]]))
				newcol = startcol + 1
				fs_results = fs_results[1:]  # Remove scale

				# Export RX data and graphs
				if excel_export_rx:  # Export the RX data and graphs the Impedance Loci
					# Insert R data in excel___________________________________________________________________________
					r_first = newcol
					r_results, x_results = [], []
					for x in fs_results:
						if x[constants.ResultsExtract.loc_pf_variable] == "m:R":
							ws.Range(ws.Cells(startrow, newcol),
									 ws.Cells(endrow, newcol)).Value = list(zip(*[x]))
							r_results.append(x[3:])
							newcol = newcol + 1
					r_last = newcol - 1

					# Insert X data in excel__________________________________________________________________________
					newcol += 1
					x_first = newcol
					for x in fs_results:
						if x[constants.ResultsExtract.loc_pf_variable] == "m:X":
							ws.Range(ws.Cells(startrow, newcol),
									 ws.Cells(endrow, newcol)).Value = list(zip(*[x]))
							x_results.append(x[3:])
							newcol = newcol + 1
					x_last = newcol - 1

					t2 = time.clock() - t1
					self.log_info('Inserting RX data self impedance data, time taken: {:.2f} seconds'.format(t2))
					t1 = time.clock()

					# Graph R X Data impedance Loci____________________________________________________________________________
					chart_width = 400  # Width of Graph
					chart_height = 300  # Height of Chart
					left = 30
					top = 45  # Top Starting Point

					if excel_export_z or excel_export_hrm:
						top = startrow * 15
					graph_across = 1  # Number of Graphs Across the Page
					graph_spacing = 25  # Spacing between the graphs

					# Adjusted to ensure that minimum number of graphs rows is 200 and not too small a number
					noofgraphrows = max(int(math.ceil(len(scale[3:]) / graph_across) - 1), 200)
					#noofgraphrowsrange = list(range(0, noofgraphrows))
					noofgraphrowsrange = list(range(0, noofgraphrows))
					gph_coord = []  # List of Graph coordinates for Impedance Loci
					for uyt in noofgraphrowsrange:  # This creates the graph coordinates
						mnb = list(range(0, graph_across))
						for lkj in mnb:
							gph_coord.append([(left + lkj * (chart_width + graph_spacing)),
											  (top + uyt * (chart_height + graph_spacing))])

					# This section is used to calculate the position of the rows for non Integer Harmonics
					scale_list_int = []
					scale_clipped = scale[3:]  # Remove Headers
					lp_count = 0
					for lkp in scale_clipped:  # Get position of harmonics
						hjg = (lkp / 50).is_integer()
						if hjg:
							scale_list_int.append(lp_count)
						# app.PrintPlain(lp_count)
						lp_count = lp_count + 1
					if len(scale_list_int) < 3:
						self.log_error("The frequency range you have given is less than 2 integer harmonics")
					else:
						diff = (scale_list_int[2] - scale_list_int[
							1]) / 2  # Get the difference between positions of whole harmonics

					non_int_rows = []
					# Plot the 1st point of range and the position of the actual harmonic and the end of harmonic
					# range [75, 100, 125] would return [0,1,2]
					for wer in scale_list_int:
						if diff < 1:
							pr = wer
							qr = wer
						else:
							pr = int(wer - diff)
							qr = int(wer + diff)
						if pr < 0:
							pr = 0
						if qr > (len(scale_clipped) - 1):
							qr = len(scale_clipped) - 1
						non_int_rows.append([pr, wer, qr])

					gc = 0  # Graph Count
					new_row = startrow + 3
					x_results = list(zip(*x_results))
					r_results = list(zip(*r_results))
					startrow1 = (endrow + 3)
					startcol1 = startcol + 3

					for hrm in non_int_rows:  # Plots the Graphs for the Harmonics including non integer rows
						ws.Range(ws.Cells(1, 1),
								 ws.Cells(1,2)).Select()  # Important for computation as it doesn't graph all the selection first ie these cells should be blank

						# Add chart - Code adjusted since don't need to select and activate each chart, can just use
						# reference
						chrt = ws.Shapes.AddChart(c.xlXYScatter, gph_coord[gc][0], gph_coord[gc][1],
												  chart_width,
												  chart_height).Chart  # AddChart(Type, Left, Top, Width, Height)
						chrt.ApplyLayout(1)  # Select Layout 1-11
						chrt.ChartTitle.Characters.Text = " Harmonic Order " + str(
							int(scale_clipped[hrm[1]] / 50))  # Add Title
						chrt.SeriesCollection(1).Delete()
						chrt.Axes(c.xlCategory).AxisTitle.Text = "Resistance in Ohms"  # X Axis
						chrt.Axes(c.xlCategory).MinimumScale = 0  # Set minimum of x axis
						chrt.Axes(c.xlCategory).TickLabels.NumberFormat = "0"  # Set number of decimals 0.0
						chrt.Axes(c.xlValue).AxisTitle.Text = "Reactance in Ohms"  # Y Axis
						chrt.Axes(c.xlValue).TickLabels.NumberFormat = "0"  # Set number of decimals

						rx_con = []

						# This is used to graph non integer harmonics on the same plot as integer
						for tres in range(hrm[0], (hrm[2] + 1)):
							# Create series to add to the plot
							series = chrt.SeriesCollection().NewSeries()
							series.XValues = ws.Range(ws.Cells((startrow + 3 + tres), r_first),
													  ws.Cells((startrow + 3 + tres), r_last))  # X Value
							series.Values = ws.Range(ws.Cells((startrow + 3 + tres), x_first),
													 ws.Cells((startrow + 3 + tres), x_last))  # Y Value
							series.Name = ws.Cells((startrow + 3 + tres), startcol)  # Name
							series.MarkerSize = 5  # Marker Size
							series.MarkerStyle = 3  # Marker type
							prop_count = 0
							if tres < len(r_results):
								for desd in r_results[tres]:
									rx_con.append([desd, x_results[tres][prop_count]])
									prop_count = prop_count + 1
							else:
								self.log_error('The Scale is longer then the dataset it probably means that you ' +
											   'have selected automatic step size adaption in FSweep')

						if excel_convex_hull:  # This is used to the convex hull of the points on the graph with a line
							rx_array = np.array(rx_con)  # Converts the RX data to a numpy array
							convex_rx = self.convex_hull(pointlist=rx_array,
														 node_name=sheet_name)  # Get the min area points of the array needs to be in numpy array
							endcol1 = (startcol1 + len(convex_rx[0]) - 1)

							# Add convex hull to excel spreadsheet
							ws.Range(ws.Cells(startrow1, startcol1),
									 ws.Cells(startrow1, endcol1)).Value = convex_rx[0]  # Adds R data to Excel
							ws.Range(ws.Cells(startrow1 + 1, startcol1),
									 ws.Cells(startrow1 + 1, endcol1)).Value = convex_rx[1]  # Add X data to Excel

							# Create new series for Convex Hull plot (could be added to a separate function)
							series = chrt.SeriesCollection().NewSeries()  # Adds a new series for it
							series.XValues = ws.Range(ws.Cells(startrow1, startcol1),
													  ws.Cells(startrow1, endcol1))  # X Value
							series.Values = ws.Range(ws.Cells(startrow1 + 1, startcol1),
													 ws.Cells(startrow1 + 1, endcol1))  # Y Value												# Name
							ws.Cells(startrow1, startcol).Value = str(int(scale_clipped[hrm[1]])) + " Hz"
							ws.Cells(startrow1, startcol + 1).Value = str(int(scale_clipped[hrm[0]])) + " Hz"
							ws.Cells(startrow1 + 1, startcol + 1).Value = str(int(scale_clipped[hrm[2]])) + " Hz"
							ws.Cells(startrow1, startcol + 2).Value = "R"
							ws.Cells(startrow1 + 1, startcol + 2).Value = "X"
							series.MarkerSize = 5  # Marker Size
							series.MarkerStyle = -4142  # Marker type
							series.Format.Line.Visible = True  # Marker line
							series.Format.Line.ForeColor.RGB = 12611584  # Colour is red + green*256 + blue*256*256
							series.Name = "Convex Hull"  # Name

							# Plots the graphs for the customers
							ws.Range(ws.Cells(1, 1), ws.Cells(1, 2)).Select()
							# Using chart reference handle rather than making active chart
							chrt = ws.Shapes.AddChart(c.xlXYScatter, gph_coord[gc][0] + 425, gph_coord[gc][1], chart_width,
												   chart_height).Chart  # AddChart(Type, Left, Top, Width, Height)
							chrt.ApplyLayout(1)  # Select Layout 1-11
							chrt.ChartTitle.Characters.Text = " Harmonic Order " + str(
								int(scale_clipped[hrm[1]] / 50))  # Add Title
							chrt.SeriesCollection(1).Delete()
							chrt.Axes(c.xlCategory).AxisTitle.Text = "Resistance in Ohms"  # X Axis
							chrt.Axes(c.xlCategory).MinimumScale = 0  # Set minimum of x axis
							chrt.Axes(
								c.xlCategory).TickLabels.NumberFormat = "0"  # Set number of decimals 0.0
							chrt.Axes(c.xlValue).AxisTitle.Text = "Reactance in Ohms"  # Y Axis
							chrt.Axes(c.xlValue).TickLabels.NumberFormat = "0"  # Set number of decimals

							# Add new series to chart
							series = chrt.SeriesCollection().NewSeries()  # Adds a new series for it
							series.XValues = ws.Range(ws.Cells(startrow1, startcol1),
													  ws.Cells(startrow1, endcol1))  # X Value
							series.Values = ws.Range(ws.Cells(startrow1 + 1, startcol1),
													 ws.Cells(startrow1 + 1, endcol1))  # Y Value												# Name
							series.Name = ws.Cells(startrow1, startcol)  # Name
							series.MarkerSize = 5  # Marker Size
							series.MarkerStyle = -4142  # Marker type
							series.Format.Line.Visible = True  # Marker line
							series.Format.Line.ForeColor.RGB = 12611584  # Colour is red + green*256 + blue*256*256

						startrow1 += 2
						new_row += 1
						gc += 1
					t2 = time.clock() - t1
					self.log_info('Graphing RX data self impedance data, time taken: {:.2f} seconds'.format(t2))

					t1 = time.clock()
					newcol = newcol + 1

				# Export Z data and graphs
				if excel_export_z:  # Export Z data and graphs
					# Insert Z data in excel_______________________________________________________________________________________________
					ws.Range(ws.Cells(startrow, newcol),
							 ws.Cells(endrow, newcol)).Value = list(zip(*[scale]))
					if excel_export_rx:
						newcol = newcol + 1
					z_first = newcol - 1
					z_results, base_case_pos = [], []
					for x in fs_results:
						if x[constants.ResultsExtract.loc_pf_variable] == "m:Z":
							ws.Range(ws.Cells(startrow, newcol), ws.Cells(endrow, newcol)).Value = list(zip(*[x]))
							z_results.append(x[3:])
							if x[constants.ResultsExtract.loc_contingency] == "Base_Case":
								base_case_pos.append([newcol])
							newcol = newcol + 1
					z_last = newcol - 1
					t2 = time.clock() - t1
					self.log_info('Inserting Z self impedance data, time taken: {:0.2f} seconds'.format(t2))

					t1 = time.clock()
					# Graph Z Data_________________________________________________________________________________________________________

					# If there is more than 1 Base Case then plot all the bases on one graph and then each base
					# against its N-1 across else just plot them all on one graph.
					if len(base_case_pos) > 1:
						z_no_of_contingencies = int(base_case_pos[1][0]) - int(base_case_pos[0][0])
						ws.Range(ws.Cells(1, 1), ws.Cells(1,
														  2)).Select()  # Important for computation as it doesn't graph all the selection first ie these cells should be blank

						# Using chart reference rather than activating chart
						chrt = ws.Shapes.AddChart(c.xlXYScatterLinesNoMarkers, 30, 45, 825, 400).Chart  # AddChart(Type, Left, Top, Width, Height)
						chrt.ApplyLayout(1)  # Select Layout 1-11
						chrt.ChartTitle.Characters.Text = sheet_name + " Base Cases m:Z Self Impedances"  # Add Title
						chrt.Axes(c.xlCategory).AxisTitle.Text = "Frequency in Hz"  # X Axis
						chrt.Axes(c.xlCategory).MinimumScale = 0  # Set minimum of x axis
						chrt.Axes(c.xlCategory).MaximumScale = scale_end  # Set maximum of x axis
						chrt.Axes(c.xlCategory).TickLabels.NumberFormat = "0"  # Set number of decimals 0.0
						chrt.Axes(c.xlValue).AxisTitle.Text = "Impedance in Ohms"  # Y Axis
						chrt.Axes(c.xlValue).TickLabels.NumberFormat = "0"  # Set number of decimals
						chrt.SeriesCollection(1).Delete()

						for zb_col in base_case_pos:
							series_name1 = ws.Range(ws.Cells((startrow + 1), zb_col[0]),
													ws.Cells((startrow + 2), zb_col[0])).Value
							series_name = str(series_name1[0][0]) + "_" + str(series_name1[1][0])
							# Using chart reference rather than active chart
							series = chrt.SeriesCollection().NewSeries()
							series.Values = ws.Range(ws.Cells((startrow + 3), zb_col[0]),
													 ws.Cells(endrow, zb_col[0]))  # Y Value
							series.XValues = ws.Range(ws.Cells((startrow + 3), z_first), ws.Cells(endrow, z_first))
							series.Name = series_name

						zb_count = 1
						for zb_col1 in base_case_pos:
							ws.Range(ws.Cells(1, 1), ws.Cells(1,
															  2)).Select()  # Important for computation as it doesn't graph all the selection first ie these cells should be blank
							# Get name of series
							series_name1 = ws.Range(ws.Cells((startrow + 1), zb_col1[0]),
													ws.Cells((startrow + 2), zb_col1[0])).Value
							series_name = str(series_name1[0][0])

							chrt = ws.Shapes.AddChart(c.xlXYScatterLinesNoMarkers, 30 + zb_count * 855, 45, 825,
													  400).Chart  # AddChart(Type, Left, Top, Width, Height)
							chrt.ApplyLayout(1)  # Select Layout 1-11
							chrt.ChartTitle.Characters.Text = sheet_name + " " + str(
								series_name) + " m:Z Self Impedances"  # Add Title
							chrt.Axes(c.xlCategory).AxisTitle.Text = "Frequency in Hz"  # X Axis
							chrt.Axes(c.xlCategory).MinimumScale = 0  # Set minimum of x axis
							chrt.Axes(c.xlCategory).MaximumScale = scale_end  # Set maximum of x axis
							chrt.Axes(
								c.xlCategory).TickLabels.NumberFormat = "0"  # Set number of decimals 0.0
							chrt.Axes(c.xlValue).AxisTitle.Text = "Impedance in Ohms"  # Y Axis
							chrt.Axes(c.xlValue).TickLabels.NumberFormat = "0"  # Set number of decimals
							chrt.SeriesCollection(1).Delete()

							# Add data series to chart
							for zzcol in list(range(zb_col1[0], (zb_col1[0] + z_no_of_contingencies))):
								series_name1 = ws.Range(ws.Cells((startrow + 1), zzcol),
														ws.Cells((startrow + 2), zzcol)).Value
								series_name = str(series_name1[0][0]) + "_" + str(series_name1[1][0])
								series = chrt.SeriesCollection().NewSeries()
								series.Values = ws.Range(ws.Cells((startrow + 3), zzcol),
														 ws.Cells(endrow, zzcol))  # Y Value
								series.XValues = ws.Range(ws.Cells((startrow + 3), z_first), ws.Cells(endrow, z_first))
								series.Name = series_name
							zb_count = zb_count + 1

					# If there is only one base case
					else:
						ws.Range(ws.Cells(startrow + 1, z_first),
								 ws.Cells(endrow, z_last)).Select()  # Important for computation as it doesn't graph all the selection first ie these cells should be blank
						
						chrt = ws.Shapes.AddChart(c.xlXYScatterLinesNoMarkers, 30, 45, 825, 400).Chart  # AddChart(Type, Left, Top, Width, Height)
						chrt.ApplyLayout(1)  # Select Layout 1-11
						chrt.ChartTitle.Characters.Text = sheet_name + " m:Z Self Impedance"  # Add Title
						chrt.Axes(c.xlCategory).AxisTitle.Text = "Frequency in Hz"  # X Axis
						chrt.Axes(c.xlCategory).MinimumScale = 0  # Set minimum of x axis
						chrt.Axes(c.xlCategory).MaximumScale = scale_end  # Set maximum of x axis
						chrt.Axes(c.xlCategory).TickLabels.NumberFormat = "0"  # Set number of decimals 0.0
						chrt.Axes(c.xlValue).AxisTitle.Text = "Impedance in Ohms"  # Y Axis
						chrt.Axes(c.xlValue).TickLabels.NumberFormat = "0"  # Set number of decimals

					t2 = time.clock() - t1
					self.log_info('Graphing Z self impedance data, time taken: {:0.2f} seconds'.format(t2))


					t1 = time.clock()

				# Export Mutual impedance data to excel
				if excel_export_z12:  # Export Z12 data
					# Insert Mutual Z_12 data to excel______________________________________________________________________________________________
					self.log_info('Inserting Z_12 data')
					res_to_include = ['c:Z_12']
					if excel_export_rx or excel_export_z:
						newcol += 1
						if excel_export_rx:
							res_to_include += ['c:R_12','c:X_12']


					# Additional loop added to loop through each string type to handle if R_12 and X_12 results
					# are exported as well.  Could be made more efficient by only looping once and separating
					# the columns by the number of results.
					for res_type in res_to_include:
						for x in fs_results:
							# #if x[1] == "c:Z_12":
							if x[constants.ResultsExtract.loc_pf_variable_mutual] == res_type:
								ws.Range(ws.Cells(startrow - 1, newcol),
										 ws.Cells(endrow, newcol)).Value = list(zip(*[x]))
								newcol = newcol + 1

						newcol = newcol + 1

					t2 = time.clock() - t1
					self.log_info('Exporting Z_12 data self impedance data, time taken: {:.2f} seconds'.format(t2))

					t1 = time.clock()

			# Save workbook so far
			wb.Save()

		# Was harmonic load flow carried out
		if hrm_sim:
			hrm_endrow = startrow + len(hrm_results[0]) - 1

			# Export harmonic data to excel
			if excel_export_hrm:
				self.log_info('Inserting Harmonic data')
				if excel_export_rx or excel_export_z or excel_export_z12:  # Adds a space between FS & harmonic data
					newcol = newcol + 1

				hrm_first = newcol
				hrm_scale = hrm_results[0]
				hrm_scale1 = [int(int(x) / 50) for x in hrm_scale[4:]]
				hrm_scale = hrm_scale[:4]
				hrm_scale.extend(hrm_scale1)
				ws.Range(ws.Cells(startrow, newcol), ws.Cells(hrm_endrow, newcol)).Value = list(
					zip(*[hrm_scale]))  # Exports the Scale to excel
				newcol += 1
				hrm_base_case_pos = []
				for x in hrm_results:  # Exports the results to excel
					if x[0] == "m:HD":
						ws.Range(ws.Cells(startrow, newcol), ws.Cells(hrm_endrow, newcol)).Value = list(zip(*[x]))
						if x[2] == "Base_Case":
							hrm_base_case_pos.append([newcol])
						newcol += 1
				hrm_last = newcol - 1

				# Graph Harmonic Distortion Charts
				if excel_export_z:
					hrm_top = 500
				else:
					hrm_top = 45

				# If there is more than 1 Base Case then plot all the bases on one graph and then each base against its N-1 across else just plot them all on one graph.
				if len(hrm_base_case_pos) > 1:
					hrm_no_of_contingencies = int(hrm_base_case_pos[1][0]) - int(hrm_base_case_pos[0][0])
					ws.Range(ws.Cells(1, 1), ws.Cells(1, 2)).Select()

					# Replaced to use chart reference rather than activating chart
					chrt = ws.Shapes.AddChart(c.xlColumnClustered, 30, hrm_top, 825, 400).Chart  # AddChart(Type, Left, Top, Width, Height)
					chrt.ApplyLayout(9)  # Select Layout 1-11
					chrt.ChartTitle.Characters.Text = sheet_name + " Base Case Harmonic Emissions v IEC Limits"  # Add Title
					chrt.SeriesCollection(1).Delete()                                     					# Delete legend
					chrt.Axes(c.xlValue).AxisTitle.Text = "HD %"  # Y Axis
					chrt.Axes(c.xlValue).TickLabels.NumberFormat = "0.0"  # Set number of decimals
					chrt.Axes(c.xlCategory).AxisTitle.Text = "Harmonic"  # X Axis
					chrt.XValues = ws.Range(ws.Cells((startrow + 3), hrm_first),
											ws.Cells(hrm_endrow, hrm_first))  # X Value

					# Add date for each harmonic result
					for hrm_col in hrm_base_case_pos:
						series = chrt.SeriesCollection().NewSeries()
						series_name1 = ws.Range(ws.Cells((startrow + 1), hrm_col[0]),
												ws.Cells((startrow + 2), hrm_col[0])).Value
						series_name = str(series_name1[0][0]) + "_" + str(series_name1[1][0])
						series.Values = ws.Range(ws.Cells((startrow + 3), hrm_col[0]),
												 ws.Cells(hrm_endrow, hrm_col[0]))  # Y Value
						series.XValues = ws.Range(ws.Cells((startrow + 3), hrm_first),
												  ws.Cells(hrm_endrow, hrm_first))
						series.Name = series_name  #
					
					# Add new series with IEC limits to datasheet and plot
					ws.Range(ws.Cells(startrow, newcol), ws.Cells(startrow + len(self.limits) - 1,
																  newcol)).Value = self.limits  # Export the limits as far as the 40th Harmonic
					series = chrt.SeriesCollection().NewSeries()  # Add series to the graph
					series.Values = ws.Range(ws.Cells(startrow + 3, newcol),
											 ws.Cells(startrow + len(self.limits) - 1, newcol))  # Y Value
					series.XValues = ws.Range(ws.Cells((startrow + 3), hrm_first), ws.Cells(hrm_endrow, hrm_first))
					series.Name = "IEC 61000-3-6"  # Name
					series.Format.Fill.Visible = True  # Add fill to chart
					series.Format.Fill.ForeColor.RGB = 12611584  # Colour for fill (red + green*256 + blue*256*256)
					series.Format.Fill.ForeColor.Brightness = 0.75  # Fill Brightness
					series.Format.Fill.Transparency = 0.75  # Fill Transparency
					# REMOVED - Statement has no effect
					# series.Format.Fill.Solid					# Solid Fill
					series.Border.Color = 12611584  # Fill Colour
					series.Format.Line.Visible = True  # Series line is visible
					series.Format.Line.Weight = 1.5  # Set line weight for series
					series.Format.Line.ForeColor.RGB = 12611584  # Colour for line (red + green*256 + blue*256*256)
					series.AxisGroup = 2  # Move to Secondary Axes
					
					# Edit chart settings to allow overlap
					# Using chart reference rather than active chart
					chrt.ChartGroups(2).Overlap = 100  # Edit Secondary Axis Overlap of bars
					chrt.ChartGroups(2).GapWidth = 0  # Edit Secondary Axis width between bars
					chrt.Axes(c.xlValue).MaximumScale = 3.5  # Set scale Max
					chrt.Axes(c.xlValue, c.xlSecondary).MaximumScale = 3.5  # Set scale Min

					hrmb_count = 1
					for hrm_col in hrm_base_case_pos:
						ws.Range(ws.Cells(1, 1), ws.Cells(1, 2)).Select()
						# Get series name
						series_name1 = ws.Range(ws.Cells((startrow + 1), hrm_col[0]),
												ws.Cells((startrow + 2), hrm_col[0])).Value
						series_name = str(series_name1[0][0])
						
						# Using chart handle rather than reference to active chart
						chrt = ws.Shapes.AddChart(c.xlColumnClustered, 30 + hrmb_count * 855, hrm_top, 825, 400).Chart  # AddChart(Type, Left, Top, Width, Height)
						chrt.ApplyLayout(9)  # Select Layout 1-11						
						chrt.ChartTitle.Characters.Text = sheet_name + " " + str(
							series_name) + " Harmonic Emissions v IEC Limits"  # Add Title
						chrt.SeriesCollection(1).Delete()
						chrt.Axes(c.xlValue).AxisTitle.Text = "HD %"  # Y Axis
						chrt.Axes(c.xlValue).TickLabels.NumberFormat = "0.0"  # Set number of decimals
						chrt.Axes(c.xlCategory).AxisTitle.Text = "Harmonic"  # X Axis
						chrt.XValues = ws.Range(ws.Cells((startrow + 3), hrm_first),
														   ws.Cells(hrm_endrow, hrm_first))  # X Value

						for hrm_col1 in list(range(hrm_col[0], (hrm_col[0] + hrm_no_of_contingencies))):
							series = chrt.SeriesCollection().NewSeries()
							series_name1 = ws.Range(ws.Cells((startrow + 1), hrm_col1),
													ws.Cells((startrow + 2), hrm_col1)).Value
							series_name = str(series_name1[0][0]) + "_" + str(series_name1[1][0])
							series.Values = ws.Range(ws.Cells((startrow + 3), hrm_col1),
													 ws.Cells(hrm_endrow, hrm_col1))  # Y Value
							series.XValues = ws.Range(ws.Cells((startrow + 3), hrm_first),
													  ws.Cells(hrm_endrow, hrm_first))
							series.Name = series_name  #
						
						# Add IEC limits to plots and excel
						ws.Range(ws.Cells(startrow, newcol), ws.Cells(startrow + len(self.limits) - 1,
																	  newcol)).Value = self.limits  # Export the limits as far as the 40th Harmonic
						series = chrt.SeriesCollection().NewSeries()  # Add series to the graph
						series.Values = ws.Range(ws.Cells(startrow + 3, newcol),
												 ws.Cells(startrow + len(self.limits) - 1, newcol))  # Y Value
						series.XValues = ws.Range(ws.Cells((startrow + 3), hrm_first),
												  ws.Cells(hrm_endrow, hrm_first))
						series.Name = "IEC 61000-3-6"  # Name
						series.Format.Fill.Visible = True  # Add fill to chart
						series.Format.Fill.ForeColor.RGB = 12611584  # Colour for fill (red + green*256 + blue*256*256)
						series.Format.Fill.ForeColor.Brightness = 0.75  # Fill Brightness
						series.Format.Fill.Transparency = 0.75  # Fill Transparency
						# REMOVED - Statement has no effect
						# series.Format.Fill.Solid														# Solid Fill
						series.Border.Color = 12611584  # Fill Colour
						series.Format.Line.Visible = True  # Series line is visible
						series.Format.Line.Weight = 1.5  # Set line weight for series
						series.Format.Line.ForeColor.RGB = 12611584  # Colour for line (red + green*256 + blue*256*256)
						series.AxisGroup = 2  # Move to Secondary Axes
						
						# Force charts to overlap
						# Use chrt reference rather than active chart
						chrt.ChartGroups(2).Overlap = 100  # Edit Secondary Axis Overlap of bars
						chrt.ChartGroups(2).GapWidth = 0  # Edit Secondary Axis width between bars
						chrt.Axes(c.xlValue).MaximumScale = 3.5  # Set scale Max
						chrt.Axes(c.xlValue, c.xlSecondary).MaximumScale = 3.5  # Set scale Min
						hrmb_count += 1
				
				# If only single base case then no need to compare base cases
				else:
					ws.Range(ws.Cells(1, 1), ws.Cells(1, 2)).Select()
					
					# Using chart reference rather than active chart
					chrt = ws.Shapes.AddChart(c.xlColumnClustered, 30, hrm_top, 825, 400).Chart  # AddChart(Type, Left, Top, Width, Height)
					chrt.ApplyLayout(9)  # Select Layout 1-11
					chrt.ChartTitle.Characters.Text = sheet_name + " Harmonic Emissions v IEC Limits"  # Add Title
					chrt.SeriesCollection(1).Delete()
					# chrt.Legend.Delete()                                                					# Delete legend
					chrt.Axes(c.xlValue).AxisTitle.Text = "HD %"  # Y Axis
					chrt.Axes(c.xlValue).TickLabels.NumberFormat = "0.0"  # Set number of decimals
					chrt.Axes(c.xlCategory).AxisTitle.Text = "Harmonic"  # X Axis
					# chrt.Axes(c.xlCategory).MinimumScale = 0                            					# Set minimum of x axis
					# chrt.Axes(c.xlCategory).TickLabels.NumberFormat = "0"               					# Set number of decimals 0.0
					chrt.XValues = ws.Range(ws.Cells((startrow + 3), hrm_first),
											ws.Cells(hrm_endrow, hrm_first))  # X Value

					for hrm_col in range(hrm_first + 1, hrm_last + 1):
						series_name1 = ws.Range(ws.Cells((startrow + 1), hrm_col),
												ws.Cells((startrow + 2), hrm_col)).Value
						series_name = str(series_name1[0][0]) + "_" + str(series_name1[1][0])
						series = chrt.SeriesCollection().NewSeries()
						series.Values = ws.Range(ws.Cells((startrow + 3), hrm_col),
												 ws.Cells(hrm_endrow, hrm_col))  # Y Value
						series.XValues = ws.Range(ws.Cells((startrow + 3), hrm_first),
												  ws.Cells(hrm_endrow, hrm_first))
						series.Name = series_name  #
					
					# Add IEC limits
					ws.Range(ws.Cells(startrow, newcol), ws.Cells(startrow + len(self.limits) - 1,
																  newcol)).Value = self.limits  # Export the limits as far as the 40th Harmonic
					series = chrt.SeriesCollection().NewSeries()  # Add series to the graph
					series.Values = ws.Range(ws.Cells(startrow + 3, newcol),
											 ws.Cells(startrow + len(self.limits) - 1, newcol))  # Y Value
					series.XValues = ws.Range(ws.Cells((startrow + 3), hrm_first), ws.Cells(hrm_endrow, hrm_first))
					series.Name = "IEC 61000-3-6"  # Name
					series.Format.Fill.Visible = True  # Add fill to chart
					series.Format.Fill.ForeColor.RGB = 12611584  # Colour for fill (red + green*256 + blue*256*256)
					series.Format.Fill.ForeColor.Brightness = 0.75  # Fill Brightness
					series.Format.Fill.Transparency = 0.75  # Fill Transparency
					# REMOVED - statement has no effect
					# series.Format.Fill.Solid														# Solid Fill
					series.Border.Color = 12611584  # Fill Colour
					series.Format.Line.Visible = True  # Series line is visible
					series.Format.Line.Weight = 1.5  # Set line weight for series
					series.Format.Line.ForeColor.RGB = 12611584  # Colour for line (red + green*256 + blue*256*256)
					series.AxisGroup = 2  # Move to Secondary Axes
					
					# Allow charts to overlap
					# Using chrt reference rather than active chart to avoid repeatedly activating the chart
					chrt.ChartGroups(2).Overlap = 100  # Edit Secondary Axis Overlap of bars
					chrt.ChartGroups(2).GapWidth = 0  # Edit Secondary Axis width between bars
					chrt.Axes(c.xlValue).MaximumScale = 3.5  # Set scale Max
					chrt.Axes(c.xlValue, c.xlSecondary).MaximumScale = 3.5  # Set scale Min

				t2 = time.clock() - t1
				self.log_info('Exporting Harmonic data, time taken: {:.2f} seconds'.format(t2))

		# Save workbook and return nothing
		wb.Save()
		return None

	def get_sheet_name(self, sheet_name, wb):
		"""
			Function checks whether the planned sheet name already exists and if it does then it returns a
			different sheet name to use when naming the worksheet
		:param str sheet_name: Planned name for worksheet
		:param wb: Handle for workbook into which new worksheet will be added
		:return str sheet_name: Worksheet name to use
		"""
		sheet_names = [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)]
		# If sheet_name is already in workbook then will need to return a new name
		if sheet_name in sheet_names:
			i = 2
			new_name = '{}({})'.format(sheet_name, i)
			# If first attempt at new_name is already in workbook then try increasing i until find one that
			# isn't already there
			while new_name in sheet_names:
				i += 1
				new_name = '{}({})'.format(sheet_name, i)
			self.log_error('Node name {} duplicated and so worksheet name {} has been used for {} instance'
						   .format(sheet_name, new_name, i))

			# Set sheet_name = new_name so that it can be returned
			sheet_name = new_name

		# Return either the new sheet_name or the original name that was used
		return sheet_name

	def convex_hull(self, pointlist, node_name):  # Gets the convex hull of a numpy array (if you have a list of tuples us np.array(pointlist) to convert
		"""
			Gets the convex hull of a numpy array
				If you have a list of tuples use np.array(pointlist) to convert
		:param np.array pointlist: Numpy array to be converted
		:param str node_name: Name of node being investigated
		:return list convex_points: List of convex points returned
		"""
		r, x = [], []
		# Potential failure here if doesn't return useful data but want script to continue with rest of results
		try:
			cv = ConvexHull(pointlist)
		except scipy.spatial.qhull.QhullError:
			self.log_error(
				'Error occurred calculating ConvexHull for {} from the following data {}'.format(node_name, pointlist))
			# Values set to 0, 0 so that something can be plotted
			err_convex_points = [[0], [0]]
			return err_convex_points

		for i in cv.vertices:
			# For each vertices extracts the R and X values
			r.append(float(pointlist[i, 0]))  # Converts the numpy floats back to regular float and attach
			x.append(float(pointlist[i, 1]))

		# Duplicates the first value of the list back to the end
		r.append(r[0])
		x.append(x[0])

		# Combine to return list containing r and x values
		convex_points = [r, x]
		return convex_points

	def close_workbook(self, wb, workbookname):  # Save and close Workbook
		"""
			Save and close the workbook
		:param Excel.Workbook wb: Handle for workbook to be closed / saved
		:param str workbookname: Full path to workbook for it to be saved as
		:return:
		"""
		self.log_info('Closing and Saving Workbook: {}'.format(workbookname))
		#SaveAs seems to throw an error so using .Save() instead since workbookname has already been set
		#wb.SaveAs(workbookname)  # Save Workbook"""
		wb.Save()
		wb.Close()  # Close Workbook
		return None

class HASTInputs:
	"""
		Class that the HAST Spreadsheet is fed into for processing
		TODO: At the moment only study settings are processed
	"""
	def __init__(self, hast_inputs, uid_time=time.strftime('%y_%m_%d_%H_%M_%S')):
		"""
			Initialises the settings based on the HAST Study Settings spreadsheet
		:param dict hast_inputs:  Dictionary of input data returned from excel_writing.Excel.import_excel_harmonic_inputs
		:param str uid_time:  Time string to use as the uid for these files
		"""
		c = constants.HASTInputs

		# Attribute definitions (study settings)
		self.pth_results_folder = str()
		self.results_name = str()
		self.progress_log_name = str()
		self.error_log_name = str()
		self.debug_log_name = str()
		self.pth_results_folder_temp = str()
		self.pf_netelm = str()
		self.pf_mutelm = str()
		self.pf_resfolder = str()
		self.pf_opscen_folder = str()
		self.pre_case_check = bool()
		self.fs_sim = bool()
		self.hrm_sim = bool()
		self.skip_failed_lf = bool()
		self.del_created_folders = bool()
		self.export_to_excel = bool()
		self.excel_visible = bool()
		self.include_rx = bool()
		self.include_convex_hull = bool()
		self.export_z = bool()
		self.export_z12 = bool()
		self.export_hrm = bool()

		# Attribute definitions (terminals)
		self.list_of_terms = list()
		self.dict_of_terms = dict()

		# Process study settings
		self.uid = uid_time
		self.study_settings(hast_inputs[c.study_settings])

		# Process List of Terminals
		self.process_terminals(hast_inputs[c.terminals])

	def study_settings(self, list_study_settings):
		"""
			Populate study settings
		:param list list_study_settings:
		:return None:
		"""
		l = list_study_settings
		# Folder to store logs (progress/error) and the excel results if empty will use current working directory
		if not l[0]:
			self.pth_results_folder = os.getcwd() + "\\"
		else:
			self.pth_results_folder = l[0]

		# Leading names to use for exported excel result file (python adds on the unique time and date).
		self.results_name = '{}{}{}.'.format(self.pth_results_folder, l[1], self.uid)
		self.progress_log_name = '{}{}{}.txt'.format(self.pth_results_folder, l[2], self.uid)
		self.error_log_name = '{}{}{}.txt'.format(self.pth_results_folder, l[3], self.uid)
		self.debug_log_name = '{}{}{}.txt'.format(self.pth_results_folder, constants.DEBUG, self.uid)

		# Temporary folder to use to store results exported during script run
		self.pth_results_folder_temp = os.path.join(self.pth_results_folder, self.uid)

		# Constants for power factory
		self.pf_netelm = l[4]
		self.pf_mutelm = '{}{}'.format(l[5], self.uid)
		self.pf_resfolder = '{}{}'.format(l[6], self.uid)
		self.pf_opscen_folder = '{}{}'.format(l[7], self.uid)

		# Constants to control study running
		self.pre_case_check = l[8]
		self.fs_sim = l[9]
		self.hrm_sim = l[10]
		self.skip_failed_lf = l[11]
		self.del_created_folders = l[12]
		self.export_to_excel = l[13]
		self.excel_visible = l[14]
		self.include_rx = l[15]
		self.include_convex_hull = l[16]
		self.export_z = l[17]
		self.export_z12 = l[18]
		self.export_hrm = l[19]

		return None

	def process_terminals(self, list_of_terminals):
		"""
			Processes the terminals from the HAST input into a dictionary so can lookup the name to use based on
			substation and terminal
		:param list list_of_terminals: List of terminals from HAST inputs, expected in the form
		:return None
		"""
		self.list_of_terms = list_of_terminals
		self.dict_of_terms = {(k[1], k[2]): k[0] for k in list_of_terminals}

		return None

	def vars_to_export(self):
		"""
			Determines the variables that will be exported from PowerFactory and they will be exported in this order
		:return list pf_vars:  Returns list of variables in the format they are defined in PowerFactory
		"""
		c = constants.PowerFactory
		pf_vars = []

		# The order variables are added here determines the order they appear in the export
		# If self impedance data should be exported
		if self.export_z:
			# Whether to include RX data as well
			if self.include_rx:
				pf_vars.append(c.pf_r1)
				pf_vars.append(c.pf_x1)
			pf_vars.append(c.pf_z1)

		# If mutual impedance data should be exported
		if self.export_z12:
			# If RX data should be exported
			if self.include_rx:
				pf_vars.append(c.pf_r12)
				pf_vars.append(c.pf_x12)
			pf_vars.append(c.pf_z12)

		return pf_vars

#  ----- UNIT TESTS -----
class TestExcelSetup(unittest.TestCase):
	"""
		UnitTest to test the operation of various excel workbook functions
	"""

	def test_excel_instance(self):
		"""
			Tests that excel instance is properly opened and closed
		"""
		with Excel(print_info=print, print_error=print) as xl:
			self.assertEqual(str(xl.xl), 'Microsoft Excel')

	def test_hast_settings_import(self):
		"""
			Tests that excel will import setting appropriately
		"""
		pth = os.path.dirname(os.path.abspath(__file__))
		pth_test_files = 'test_file_store'
		test_workbook = 'HAST_test_inputs.xlsx'
		input_file = os.path.join(pth, pth_test_files, test_workbook)
		with Excel(print_info=print, print_error=print) as xl:
			analysis_dict = xl.import_excel_harmonic_inputs(workbookname=input_file)
			self.assertEqual(len(analysis_dict.keys()), 7)

	def test_create_close_workbook(self):
		"""
			Tests that excel will create a new workbook appropriately
		"""
		pth = os.path.dirname(os.path.abspath(__file__))
		pth_test_files = 'test_file_store'
		test_workbook = 'HAST_test_outputs.xlsx'
		output_file = os.path.join(pth, pth_test_files, test_workbook)
		with Excel(print_info=print, print_error=print) as xl:
			wb = xl.create_workbook(workbookname=output_file, excel_visible=False)
			self.assertTrue(os.path.isfile(output_file))

			xl.close_workbook(wb=wb, workbookname=output_file)
			os.remove(output_file)
			self.assertFalse(os.path.isfile(output_file))

	def test_sheet_name(self):
		"""
			Tests checking whether a worksheet name already exists
		"""
		# Create unnamed workbook
		with Excel(print_info=print, print_error=print) as xl:
			wb = xl.xl.Workbooks.Add()
			sht_name = 'Sheet1'
			# Confirm that the returned value does not equal the provided value
			self.assertFalse(sht_name==xl.get_sheet_name(sht_name, wb))
			wb.Close()


