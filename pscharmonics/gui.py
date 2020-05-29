"""
	Script to handle production of GUI and returning files to be processed together to user

	TODO: Main GUI needs to have
	TODO 1. Button for study file
	TODO 2. Button for review / edit settings (pop-up with another window for each of the settings files being used)
	TODO 3. Button for pre-case check runner
	TODO 4. Button for pre-case check status (click to open results)
	TODO 5. Button to run studies
"""
import tkinter as tk
import tkinter.filedialog
import tkinter.scrolledtext
import tkinter.messagebox as messagebox
import tkinter.ttk as ttk
import sys
import os
import logging
import webbrowser
from PIL import Image, ImageTk

import pscharmonics
import pscharmonics.constants as constants
import pscharmonics.file_io as file_io
import inspect

logger = logging.getLogger()

def file_selector(initial_pth='', open_file=False, save_dir=False,
				  save_file=False,
				  lbl_file_select='Select results file(s) to add',
				  lbl_folder_select='Select folder to store results in',
				  def_ext=constants.Results.extension,
				  openfile_types=(('CSV files', '*.csv'),
							  ('All Files', '*.*'))
				  ):
	"""
		Function to allow the user to select a file to either open or save
	:param str initial_pth: (optional='') Path to use as starting location
	:param bool open_file: (optional=True) will ask for a file to open
	:param bool save_dir: (optional=False) will ask for a directory into which the results should be saved
	:param bool save_file: (optional=False) will ask for a file name to save the results as
	:param str lbl_file_select: (optional) = Title for dialog box that pops up and asks for user to select results file
	:param str lbl_folder_select: (optional) = Title for dialog box when asking user to select a folder
	:param str def_ext: (optional) = Default extension for results export
	:param tuple openfile_types: (optional) = Extension options for selecting a file type to open
	:return list file_paths:  List of paths or if target folder returned then this is the only item in the list
	"""
	file_paths = ['']
	# Determine initial_pth if not provided
	if initial_pth == '':
		# If no path provided then start with path of current script
		initial_pth, _ = os.path.split(sys.argv[0])


	# Determine whether a open_file or save_file request
	if all([open_file, save_dir]):
		raise SyntaxError('Attempted to get both an open and save file')
	elif not any([open_file, save_dir, save_file]):
		raise SyntaxError('No open or save statement provided')

	# Load tkinter window and then hide
	root = tk.Tk()
	root.withdraw()

	if open_file:
		# Load window asking user to select files for import
		_files = tkinter.filedialog.askopenfilenames(initialdir=initial_pth,
											   title=lbl_file_select,
											   filetypes=openfile_types)
		file_paths = _files
		root.destroy()
	elif save_file:
		# Load window asking user to provide file name to save extracted results as
		_file = tkinter.filedialog.asksaveasfilename(initialdir=initial_pth,
													 title=lbl_file_select,
													 filetypes=((def_ext,
																 '*{}'.format(def_ext)),
																('All Files', '*.*'))
													 )
		root.destroy()
		file_paths = [_file]
	elif save_dir:
		# Load window asking user to select destination in which to save results
		_file = tkinter.filedialog.askdirectory(initialdir=initial_pth,
										  title=lbl_folder_select)
		root.destroy()
		file_paths = [_file]

	return file_paths


class CustomStyles:
	""" Class used to customize the layout of the GUI """
	def __init__(self):
		"""
			Initialise the reference to the style names
		"""
				# Constants for styles
		# Style for Loading the SAV Button
		self.cmd_buttons = 'General.TButton'
		self.label_general = 'TLabel'
		self.label_mainheading = 'MainHeading.TLabel'
		self.label_version_number = 'Version.TLabel'
		self.label_psc_info = 'PSCInfo.TLabel'
		self.label_psc_phone = 'PSCPhone.TLabel'
		self.label_hyperlink = 'Hyperlink.TLabel'

		self.configure_styles()

	def configure_styles(self):
		"""
			Function configures all the ttk styles used within the GUI
			Further details here:  https://anzeljg.github.io/rin2/book2/2405/docs/tkinter/ttk-style-layer.html
		:return:
		"""
		# Tidy up the repeat ttk.Style() calls
		# Switch to a different theme
		styles = ttk.Style()
		styles.theme_use('clam')

		# Configure the same font in all labels
		standard_font = constants.GuiDefaults.font_family
		bg_color = constants.GuiDefaults.color_main_window

		s = ttk.Style()
		s.configure('.', font=(standard_font, '8'))

		# General style for all buttons and active color changes
		s.configure(self.cmd_buttons, height=2, width=25)

		s.configure(self.label_general, background=bg_color)
		s.configure(self.label_mainheading, font=(standard_font, '10', 'bold'), background=bg_color,
					foreground=constants.GuiDefaults.font_heading_color)
		s.configure(self.label_version_number, font=(standard_font, '7'), background=bg_color, justify=tk.CENTER)
		s.configure(self.label_hyperlink, foreground='Blue', font=(standard_font, '7'), justify=tk.CENTER)

		s.configure(
			self.label_psc_info, font=constants.GuiDefaults.psc_font,
			color=constants.GuiDefaults.psc_color_web_blue, justify='center', background=bg_color
		)

		s.configure(
			self.label_psc_phone, font=(constants.GuiDefaults.psc_font, '8'),
			color=constants.GuiDefaults.psc_color_grey, background=bg_color
		)

		return None

	def command_button_color_change(self, color):
		"""
			Force change in command button color to highlight error
		"""
		s = ttk.Style()
		s.configure(self.cmd_buttons, background=color)

		return None

class MainGui:
	"""
		Main class to produce the GUI for user interaction
		Allows the user to set up the parameters and define the cases to run the studies

	"""
	inputs = None # type: file_io.StudyInputs
	pf = None # type: pf.PowerFactory
	pf_projects = None # type: dict


	def __init__(
			self, title=constants.GuiDefaults.gui_title, start_directory=os.path.dirname(os.path.realpath(__file__))
	):
		"""
		Initialise GUI
		:param title: (optional) - Title to be used for main window

		"""
		self.logger = logging.getLogger(constants.logger_name)

		# Constants defined later
		self.pre_case_file = str()
		self.results_file = str()


		# Initial directory that will be used whenever a file selection is necessary
		self.init_dir = start_directory
		# Status set to True if user aborts rather than running studies
		self.abort = False

		# Initialise constants and Tk window
		self.master = tk.Tk()
		self.master.title(title)

		# Change color of main window
		self.master.configure(bg=constants.GuiDefaults.color_main_window)

		# Ensure that on_closing is processed correctly
		self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

		# General constants which needs to be initialised
		self._row = 0
		self._col = 0

		# Configure styles
		self.styles = CustomStyles()

		# Initialisation
		# Add GUI title
		_ = self.add_main_label(row=self.row(1), col=self.col())

		# Add button for selecting Settings file
		self.button_select_settings = self.add_cmd(
			label=constants.GuiDefaults.button_select_settings_label,
			cmd=self.load_settings_file, tooltip='Click to select the file which contains the study settings to be run'
		)
		self.lbl_settings_file = self.add_minor_label(row=self.row(), col=self.col()+1, label='No settings file selected')

		self.lbl_status = self.add_minor_label(row=self.row(1), col=self.col(), label='')

		# Add button to run a pre-case check
		self.button_precase_check = self.add_cmd(
			label='Run Pre-case Check',	cmd=self.run_precase_check,
			tooltip='Click to run a pre-case check (optional)', state=tk.DISABLED
		)

		# Add button for pre_case check results
		self.button_precase_results = self.add_cmd(
			label='Review Pre-Case Check Results',	cmd=lambda results='pre': self.load_results(results=results),
			tooltip='Click to review the pre-case results in excel', state=tk.DISABLED,
			# Row and column numbers declared, aligns with run pre-case check but offset but 1
			row=self.row(), col=self.col()+1
		)

		# Add button to run and review studies
		self.button_run_studies = self.add_cmd(
			label='Run Studies',	cmd=self.run_studies,
			tooltip='Click to run studies', state=tk.DISABLED
		)
		# Add button for pre_case check results
		self.button_study_results = self.add_cmd(
			label='Review Study Results',	cmd=lambda results='post': self.load_results(results=results),
			tooltip='Click to review the study results', state=tk.DISABLED,
			# Row and column numbers declared, aligns with run pre-case check but offset but 1
			row=self.row(), col=self.col()+1
		)

		# Separator
		self.add_sep(row=self.row(1), col_span=2)
		_ = self.add_main_label(row=self.row(1), col=self.col(), label='Combine Previous Results')

		# Add button to combine previous sets of results and produce loci
		self.button_run_previous_results = self.add_cmd(
			label='Combine Previous Results',	cmd=self.combine_results,
			tooltip='Combine previously run results into a single excel spreadsheet', state=tk.NORMAL
		)

		# Add button for pre_case check results
		self.previous_results = self.add_cmd(
			label='Review Combined Study Results',	cmd=lambda results='post': self.load_results(results=results),
			tooltip='Click to open excel with the combined study results', state=tk.DISABLED,
			row=self.row(), col=self.col()+1
		)

		# Separator before PSC details
		self.add_sep(row=self.row(1), col_span=2)

		# Add PSC logo in Windows Manager
		self.add_psc_logo_wm()

		# Add PSC logo with hyperlink to the website
		self.add_logo(
			row=self.row(1), col=self.col(),
			img_pth=constants.GuiDefaults.img_pth_psc_main,
			hyperlink=constants.GuiDefaults.hyperlink_psc_website,
			tooltip='Clicking will open PSC website',
			size=constants.GuiDefaults.img_size_psc
		)

		# Add link to the user manual and reference to the tool version
		self.add_hyp_user_manual(row=self.row(1), col=self.col())

		# Buttons that should be enabled once the inputs are loaded
		self.buttons_to_enable_level1 = (
			self.button_precase_check,
			self.button_run_studies
		)

		self.logger.debug('GUI window created')

		# to make sure GUI open in front of PSSE gui
		# self.master.deiconify()
		# Produce GUI window
		self.master.mainloop()

	def row(self, i=0):
		"""
			Returns the current row number + i
		:param int i: (optional=0) - Will return the current row number + this value
		:return int _row:
		"""
		self._row += i
		return self._row

	def col(self, i=0):
		"""
			Returns the current col number + i
		:param int i: (optional=0) - Will return the current col number + this value
		:return int _row:
		"""
		self._col += i
		return self._col

	def add_sep(self, row, col_span):
		"""
			Function just adds a horizontal separator
		:param row: Row number to use
		:param col_span: Column span number to use
		:return None:
		"""
		# Add separator
		sep = ttk.Separator(self.master, orient="horizontal")
		sep.grid(row=row, sticky=tk.W + tk.E, columnspan=col_span, pady=10)
		return None

	def add_psc_logo_wm(self):
		"""
			Function just adds the PSC logo to the windows manager in GUI
		:return: None
		"""
		# Create the PSC logo for including in the windows manager
		logo = tk.PhotoImage(file=constants.GuiDefaults.img_pth_psc_window)
		# noinspection PyProtectedMember
		self.master.tk.call('wm', 'iconphoto', self.master._w, logo)
		return None

	def add_logo(self, row, col, img_pth, hyperlink=None, tooltip=None, size=constants.GuiDefaults.img_size_psc):
		"""
			Function to add an image which when clicked is a hyperlink to the companies logo.
			Image is added using a label and changing the it to be an image and binding a hyperlink to it
		:param int row:  Row number to use
		:param int col:  Column number to use
		:param str img_pth:  Path to image to use
		:param str hyperlink:  (optional=None) Website hyperlink to use
		:param str tooltip:  (Optional=None) Popup message to use for mouse over
		:param tuple size: (optional) - Size to make image when inserting
		:return ttk.Label logo:  Reference to the newly created logo
		"""
		# Load the image and convert into a suitable size for displaying on the GUI
		img = Image.open(img_pth)
		img.thumbnail(size)
		# Convert to a photo image for inclusion on the GUI
		img_to_include = ImageTk.PhotoImage(img)

		# Add image to GUI
		logo = tk.Label(self.master, image=img_to_include, cursor='hand2', justify=tk.CENTER, compound=tk.TOP, bg='white')
		logo.photo = img_to_include
		logo.grid(row=row, column=col, columnspan=2, pady=10)

		# Associate clicking of the button as opening a web browser with the provided hyperlink
		if hyperlink:
			logo.bind(
				constants.GuiDefaults.mouse_button_1,
				lambda e: webbrowser.open_new(hyperlink)
			)

		# Add tooltip for hovering over button
		CreateToolTip(widget=logo, text=tooltip)

		return logo

	def add_hyp_user_manual(self, row, col):
		"""
			Function just adds the version and hyperlink to the user manual to the GUI
		:param row: Row Number to use
		:param col: Column number to use
		:return None:
		"""
		version_tool = ttk.Label(
			self.master, text='Version: {}'.format(constants.__version__),
			style=self.styles.label_version_number
		)
		version_tool.grid(row=row, column=col, padx=5, pady=5)

		# Create user manual link and reference to the version of the tool
		hyp_user_manual = ttk.Label(
			self.master, text="User Guide", cursor="hand2", style=self.styles.label_hyperlink
		)
		hyp_user_manual.grid(row=row, column=col + 1, padx=5, pady=5)
		hyp_user_manual.bind(constants.GuiDefaults.mouse_button_1, lambda e: webbrowser.open_new(
			os.path.join(constants.local_directory, constants.General.user_guide_reference)))

		CreateToolTip(widget=hyp_user_manual, text=(
			"Open the GUI user guide"
		))
		return None

	def add_main_label(self, row, col, label=constants.GuiDefaults.gui_title):
		"""
			Function to add the name to the GUI
		:param row: Row number to use
		:param col: Column number to use
		:param str label: (optional) = Label to use for header
		:return ttk.Label lbl:  Reference to the newly created label
		"""
		# Add label with the name to the GUI
		lbl = ttk.Label(self.master, text=label, style=self.styles.label_mainheading)
		lbl.grid(row=row, column=col, columnspan=2, pady=5, padx=10)
		return lbl

	def add_minor_label(self, row, col, label):
		"""
			Function to add the name to the GUI
		:param row: Row number to use
		:param col: Column number to use
		:param str label: (optional) = Label to use for header
		:return ttk.Label lbl:  Reference to the newly created label
		"""
		# Add label with the name to the GUI
		lbl = ttk.Label(self.master, text=label, style=self.styles.label_general)
		lbl.grid(row=row, column=col, columnspan=2, pady=5)
		return lbl

	def add_cmd(self, label, cmd, state=tk.NORMAL, tooltip=str(), row=None, col=None):
		"""
			Function just adds the command button to the GUI which is used for loading the SAV case
		:param int row: (optional) Row number to use
		:param int col: (optional) Column number to use
		:param str label:  Label to use for button
		:param func cmd: Command to use when button is clicked
		:param int state:  Tkinter state for button initially
		:param str tooltip:  Message that pops up if hover over button
		:return None:
		"""
		# If no number is provided for row or column then assume to add 1 to row and 0 to column
		if not row:
			row = self.row(1)
		if not col:
			col = self.col()

		button = ttk.Button(
			self.master, text=label, command=cmd, style=self.styles.cmd_buttons, state=state)
		button.grid(row=row, column=col, padx=5, pady=5, sticky=tk.W+tk.E)
		CreateToolTip(widget=button, text=tooltip)

		return button

	def review_edit_settings(self):
		"""
			Function to load another window which allows the user to review / edit the PowerFactory
			settings which have been provided as inputs
		:return:
		"""
		# Warning message until function fully implemented
		frame = inspect.currentframe()
		self.logger.warning('Function <{}> not yet implemented'.format(inspect.getframeinfo(frame).function))

	def run_precase_check(self):
		"""
			Function to run a pre-case check on all the study files and then provide the option for the user to
			view any issues with the loaded study cases
		:return:
		"""

		# Ask user for file to save results of pre_case check into
		pth_precase = tk.filedialog.asksaveasfilename(
			initialdir=self.init_dir,
			initialfile='Pre Case Check_{}.xlsx'.format(constants.uid),
			filetypes=constants.GuiDefaults.xlsx_types,
			title='Select the file to save the results of the pre-case check into'
		)

		if pth_precase:
			self.pre_case_file = pth_precase

			# Run the pre-case check
			_ = pscharmonics.pf.run_pre_case_checks(
				pf_projects=self.pf_projects,
				terminals=self.inputs.terminals,
				include_mutual=self.inputs.settings.export_mutual,
				export_pth=self.pre_case_file,
				contingencies=self.inputs.contingencies,
				contingencies_cmd=self.inputs.contingency_cmd
			)

			# Needs to enable the precase check button
			self.button_precase_results.configure(state=tk.NORMAL)
		else:
			self.logger.warning('No pre-case results file selected')

		return None

	def combine_results(self):
		"""
			Function to ask the user to select previous results and combine into a single results file
		:return:
		"""
		raise SyntaxError('Not developed yet')

		# # Ask user for file to save results of pre_case check into
		# pth_precase = tk.filedialog.asksaveasfilename(
		# 	initialdir=self.init_dir,
		# 	initialfile='Pre Case Check_{}.xlsx'.format(constants.uid),
		# 	filetypes=constants.GuiDefaults.xlsx_types,
		# 	title='Select the file to save the results of the pre-case check into'
		# )
		#
		# if pth_precase:
		# 	self.pre_case_file = pth_precase
		#
		# 	# Run the pre-case check
		# 	_ = pscharmonics.pf.run_pre_case_checks(
		# 		pf_projects=self.pf_projects,
		# 		terminals=self.inputs.terminals,
		# 		include_mutual=self.inputs.settings.export_mutual,
		# 		export_pth=self.pre_case_file,
		# 		contingencies=self.inputs.contingencies,
		# 		contingencies_cmd=self.inputs.contingency_cmd
		# 	)
		#
		# 	# Needs to enable the precase check button
		# 	self.button_precase_results.configure(state=tk.NORMAL)
		# else:
		# 	self.logger.warning('No pre-case results file selected')
		#
		# return None

	def load_results(self, results):
		"""
			Loads a spreadsheet with the pre-case check results
		:param str results: ('pre' = Pre-case, 'post' = Final)
		:return None:
		"""

		if results=='pre':
			if os.path.isfile(self.pre_case_file):
				# Launch excel with the pre_case file open
				os.system('start excel.exe "%s"' % (self.pre_case_file, ))
			else:
				self.logger.critical('No file has been created at the target path {}'.format(self.pre_case_file))
				raise RuntimeError('Error running the pre-case checks')

		elif results=='post':
			if os.path.isfile(self.pre_case_file):
				# Launch excel with the pre_case file open
				os.system('start excel.exe "%s"' % (self.results_file, ))
			else:
				self.logger.critical('No file has been created at the target path {}'.format(self.results_file))
				raise RuntimeError('Error running the pre-case checks')

		else:
			raise SyntaxError('An error occurred and the load_results method has been passed the wrong sort of input')

		return None

	def run_studies(self):
		"""
			Runs the full studies
		:return:
		"""
		# Ask user for file to save results of pre_case check into
		pth_results = tk.filedialog.asksaveasfilename(
			initialdir=self.init_dir,
			initialfile='Results_{}.xlsx'.format(constants.uid),
			filetypes=constants.GuiDefaults.xlsx_types,
			title='Select the file to save the overall results to'
		)

		if pth_results:
			self.results_file = pth_results

			# Set the export folder for the inputs to be a new folder with the same name as the pth_results
			self.inputs.settings.add_folder(pth_results)


			# Run the pre-case check
			try:
				_ = pscharmonics.pf.run_studies(
					pf_projects=self.pf_projects,
					inputs=self.inputs
				)

				_ = file_io.ExtractResults(target_file=self.results_file, search_pths=(self.inputs.settings.export_folder, ))

				# Needs to enable the results check button
				self.button_study_results.configure(state=tk.NORMAL)

			except RuntimeError:
				self.lbl_status.configure(
					text='ERROR: Unable to run studies, could be a license issue, check the error messages!'
				)

				self.styles.command_button_color_change(color=constants.GuiDefaults.error_color)


		else:
			self.logger.warning('No results file selected')

		return None

	def load_settings_file(self):
		"""
			Function to allow the user to select the settings file which then once imported enables further buttons
			The function to run the settings file is housed elsewhere
		:return None:
		"""
		# Minimise main window until settings file is loaded
		# self.master.iconify()

		# Ask user to select file(s) or folders
		pth_settings = tk.filedialog.askopenfilename(
			initialdir=self.init_dir,
			filetypes=constants.GuiDefaults.xlsx_types,
			title='Select the PSC Harmonics input spreadsheet to use for this study'
		)

		if not os.path.isfile(pth_settings):
			# If the file is empty or not a genuine file then log a message and return
			self.logger.warning('File {} not found, please select a different file'.format(pth_settings))
		else:
			# Import the settings file and check if import successful
			self.lbl_status.configure(text='Loading settings file')
			self.master.update()
			self.inputs = file_io.StudyInputs(pth_file=pth_settings, gui_mode=True)

			if self.inputs.error:
				# If there is an error when importing workbook
				self.logger.error(
					(
						'Error when trying to import workbook: {}, see messages above and either select a different '
						'workbook or correct this one'
					).format(pth_settings)
				)
				self.lbl_settings_file.configure(text='Settings file error')

			else:
				# Update the path and name of the file
				self.init_dir, file_name = os.path.split(pth_settings)

				# Initialise PowerFactory and associated projects
				self.intialise_pf_and_load_projects()

				if self.pf_projects:
					# If importing workbook was a success then change the state of the other buttons and continues
					for button in self.buttons_to_enable_level1:
						button.configure(state=tk.NORMAL)

					# Change label for button to reference the study settings
					self.lbl_settings_file.configure(text='Settings file: {} loaded'.format(file_name))
				else:
					self.lbl_settings_file.configure(
						text='Settings file: {} loaded but unable to initialise projects'.format(file_name)
					)

		# Return the parent window
		# self.master.deiconify()

		return None

	def intialise_pf_and_load_projects(self):
		"""
			Function will carry out the initial loading of PowerFactory and then create references to all the projects
		:return None:
		"""
		self.logger.info('Initialising PowerFactory')

		# Initialise PowerFactory
		self.lbl_status.configure(text='Initialising PowerFactory and projects...')
		self.master.update()
		self.pf = pscharmonics.pf.PowerFactory()
		self.pf.initialise_power_factory()

		if self.pf.pf_initialised:
			msg = 'PowerFactory Initialised, loading cases'
			self.logger.info(msg)
		else:
			msg = 'ERROR: Failed to initialise PowerFactory'
			self.logger.error(msg)
		self.lbl_status.configure(text=msg)


		self.logger.debug('Initialising PowerFactory project instances')
		self.pf_projects = pscharmonics.pf.create_pf_project_instances(
			df_study_cases=self.inputs.cases,
			lf_settings=self.inputs.lf_settings,
			fs_settings=self.inputs.fs_settings
		)

		if self.pf_projects:
			msg = 'PowerFactory projects initialised'
			self.logger.info(msg)
		else:
			msg = 'ERROR: Failed to initialise PowerFactory projects'
			self.logger.error(msg)
		self.lbl_status.configure(text=msg)

		return None

	def on_closing(self):
		"""
			Function runs when window is closed to determine if user actually wants to cancel running of study
		:return None:
		"""
		# Ask user to confirm that they actually want to close the window
		result = messagebox.askquestion(
			title='Exit study?',
			message='Are you sure you want to stop this study?',
			icon='warning'
		)

		# Test what option the user provided
		if result == 'yes':
			if self.pf_projects:
				# Delete the temporary folders created for each project if required as part of the input settings
				if self.inputs.settings.delete_created_folders:
					self.logger.debug('Early closure of GUI so deleting temporarily created folders')
					for project_name, pf_project in self.pf_projects:
						pf_project.delete_temp_folders()
				else:
					self.logger.debug('Early closure of GUI but no folders created')

			# Close window
			self.abort = True
			self.master.destroy()
			return None
		else:
			return None

class CreateToolTip(object):
	"""
		Function to create a popup tool tip for a given widget based on the descriptions provided here:
			https://stackoverflow.com/questions/3221956/how-do-i-display-tooltips-in-tkinter
	"""

	def __init__(self, widget, text="widget info"):
		"""
			Establish link with tooltip
		:param widget: Tkinter element that tooltip should be associated with
		:param text:    Message to display when hovering over button
		"""
		self.wait_time = 500  # milliseconds
		self.wrap_length = 450  # pixels
		self.widget = widget
		self.text = text
		self.widget.bind("<Enter>", self.enter)
		self.widget.bind("<Leave>", self.leave)
		self.widget.bind("<ButtonPress>", self.leave)
		self.id = None
		self.tw = None

	def enter(self, event=None):
		del event
		self.schedule()

	def leave(self, event=None):
		del event
		self.unschedule()
		self.hidetip()

	def schedule(self, event=None):
		del event
		self.unschedule()
		self.id = self.widget.after(self.wait_time, self.showtip)

	def unschedule(self, event=None):
		del event
		_id = self.id
		self.id = None
		if _id:
			self.widget.after_cancel(_id)

	def showtip(self):
		x, y, cx, cy = self.widget.bbox("insert")
		x += self.widget.winfo_rootx() + 25
		y += self.widget.winfo_rooty() + 20
		# creates a top level window
		self.tw = tk.Toplevel(self.widget)
		self.tw.attributes('-topmost', 'true')
		# Leaves only the label and removes the app window
		self.tw.wm_overrideredirect(True)
		self.tw.wm_geometry("+%d+%d" % (x, y))
		label = tk.Label(
			self.tw, text=self.text, justify='left', background="#ffffff", relief='solid', borderwidth=1,
			wraplength=self.wrap_length
		)
		label.pack(ipadx=1)

	def hidetip(self):
		tw = self.tw
		self.tw = None
		if tw:
			tw.destroy()
