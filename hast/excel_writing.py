"""
	New script created to deal with the reading and writing of data to excel to try and improve processing
	performance
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
		self.analysis_sheets = (
		("Study_Settings", "B5"), ("Base_Scenarios", "A5"), ("Contingencies", "A5"), ("Terminals", "A5"),
		("Loadflow_Settings", "D5"), ("Frequency_Sweep", "D5"), ("Harmonic_Loadflow", "D5"))
		# IEC limits
		# TODO: Should be moved to a constants or imported from inputs workbook
		# If on input spreadsheet then can be used to test against allocated limits
		self.iec_limits = [
			["IEC", "61000-3-6", "Harmonics", "THD", 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18,
			 19, 20,
			 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40],
			["IEC", "61000-3-6", "Limits", 3, 1.4, 2, 0.8, 2, 0.4, 2, 0.4, 1, 0.35, 1.5, 0.32, 1.5, 0.3, 0.3,
			 0.28, 1.2, 0.265, 0.93, 0.255, 0.2, 0.246, 0.88,
			 0.24, 0.816, 0.233, 0.2, 0.227, 0.703, 0.223, 0.66, 0.219, 0.2, 0.2158, 0.58, 0.2127, 0.55, 0.21,
			 0.2, 0.2075]]
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
			# Code only to be executed for these sheets
			# TODO: 'Contingencies', 'Base_Scenarios' and 'Terminals' should be defined in <constants.py>
			if x[0] in ('Contingencies', 'Base_Scenarios', 'Terminals'):
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

					# TODO: Rewrite as function
					# Routine only for 'Contingencies' worksheet
					if x[0] == "Contingencies":
						if len(row_value) > 2:
							aa = row_value[1:]
							station_name = aa[0::3]
							breaker_name = aa[1::3]
							breaker_status = aa[2::3]
							breaker_name1 = [str(nam) + ".ElmCoup" for nam in breaker_name]
							aa1 = list(zip(station_name, breaker_name1, breaker_status))
							aa1.insert(0, row_value[0])
						else:
							aa1 = [row_value[0], [0]]
						row_value = aa1

					# Routine for Base_Scenarios worksheet
					elif x[0] == "Base_Scenarios":
						row_value = [
							row_value[0],
							row_value[1],
							'{}.IntCase'.format(row_value[2]),
							'{}.IntScenario'.format(row_value[3])]

					# Routine for Terminals worksheet
					elif x[0] == "Terminals":
						row_value = [
							row_value[0],
							'{}.ElmSubstat'.format(row_value[1]),
							'{}.ElmTerm'.format(row_value[2])]

					row_input.append(row_value)
					count_row = count_row + 1

			# More efficiently checking which worksheet looking at
			elif x[0] in ('Study_Settings', 'Loadflow_Settings', 'Frequency_Sweep', 'Harmonic_Loadflow'):
				row_value = ws.Range(cell_start + ":" + row_end).Value
				for item in row_value:
					row_input.append(item[0])
				if x[0] == "Loadflow_Settings":
					# Process inputs for Loadflow_Settings
					z = row_input
					row_input = [
						int(z[0]), int(z[1]), int(z[2]), int(z[3]), int(z[4]), int(z[5]), int(z[6]), int(z[7]),
						int(z[8]),
						float(z[9]), int(z[10]), int(z[11]), int(z[12]), z[13], z[14], int(z[15]), int(z[16]),
						int(z[17]),
						int(z[18]), float(z[19]), int(z[20]), float(z[21]), int(z[22]), int(z[23]), int(z[24]),
						int(z[25]),
						int(z[26]), int(z[27]), int(z[28]), z[29], z[30], int(z[31]), z[32], int(z[33]), int(z[34]),
						int(z[35]), int(z[36]), int(z[37]), z[38], z[39], z[40], z[41], int(z[42]), z[43], z[44], z[45],
						z[46], z[47], z[48], z[49], z[50], z[51], int(z[52]), int(z[53]), int(z[54])]

				elif x[0] == "Frequency_Sweep":
					# Process inputs for Frequency_Sweep settings
					z = row_input
					row_input = [z[0], z[1], int(z[2]), z[3], z[4], z[5], int(z[6]), z[7], z[8], z[9],
								 z[10], z[11], z[12], z[13], int(z[14]), int(z[15])]

				elif x[0] == "Harmonic_Loadflow":
					# Process inputs for Harmonic_Loadflow
					z = row_input
					row_input = [int(z[0]), int(z[1]), int(z[2]), int(z[3]), z[4], z[5], z[6], z[7],
								 z[8], int(z[9]), int(z[10]), int(z[11]), int(z[12]), int(z[13]), int(z[14])]

			# Combine imported results in a dictionary relevant to the worksheet that has been imported
			analysis_dict[(x[0])] = row_input  # Imports range of values into a list of lists

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
						if x[0] == "m:R":
							ws.Range(ws.Cells(startrow, newcol),
									 ws.Cells(endrow, newcol)).Value = list(zip(*[x]))
							r_results.append(x[3:])
							newcol = newcol + 1
					r_last = newcol - 1

					# Insert X data in excel__________________________________________________________________________
					newcol += 1
					x_first = newcol
					for x in fs_results:
						if x[0] == "m:X":
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

					noofgraphrows = int(math.ceil(len(scale[3:]) / graph_across) - 1)
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
						if x[0] == "m:Z":
							ws.Range(ws.Cells(startrow, newcol), ws.Cells(endrow, newcol)).Value = list(zip(*[x]))
							z_results.append(x[3:])
							if x[2] == "Base_Case":
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
					if excel_export_rx or excel_export_z:
						newcol += 1

					for x in fs_results:
						if x[1] == "c:Z_12":
							ws.Range(ws.Cells(startrow - 1, newcol),
									 ws.Cells(endrow, newcol)).Value = list(zip(*[x]))
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


