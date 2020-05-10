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


class MainGUI:
	"""
		Main class to produce and store the GUI
		Allows the user to select files to be processed and displays a list of all the files which will be processed
		Once the user has selected some files it will enable options to select the filtering parameters for the
		results
	"""
	def __init__(self, title='Results Processing', start_directory='', files=False,
				 def_ext=constants.Results.extension):
		"""
			Initialise GUI
		:param str title: (optional) - Title to be used for main window
		:param str start_directory: (optional) - Path to a directory to use for processing
		:param bool files: (optional) = False, when set to False will ask user to select Folders rather than files
		:param str def_ext: (optional) = '.xlsx', extension of files to search for when using a folder selection
		"""
		# General constants which need to be initialised
		self._row = 0
		self._col = 0

		# Initial results pth assumed to be same as script location and is then updated for when each file is selected
		if start_directory == '':
			self.results_pth = os.path.dirname(os.path.abspath(__file__))
		else:
			self.results_pth = start_directory

		# Is populated with a list of file paths to be returned
		self.results_files_list = []
		# Target file to export results to
		self.target_file = ''

		# Initialise constants and tk window
		self.master = tk.Tk()
		self.master.title = title
		self.bo_files = files
		self.ext = def_ext

		# Add command button for user to select files / folders
		# if files=True then ask user to select file, if files = False
		if self.bo_files:
			lbl_button = 'Add File'
		else:
			lbl_button = 'Add Folder'
		self.cmd_add_file = tk.Button(self.master,
										   text=lbl_button,
										   command=self.add_new_file)
		self.cmd_add_file.grid(row=self.row(), column=self.col())

		_ = tk.Label(master=self.master,
						  text='Files to be compared:')
		_.grid(row=self.row(1), column=self.col())

		self.lbl_results_files = tkinter.scrolledtext.ScrolledText(master=self.master)
		self.lbl_results_files.grid(row=self.row(1), column=self.col())
		self.lbl_results_files.insert(tk.INSERT,
									  'No Folders Selected')

		self.cmd_process_results = tk.Button(self.master,
												  text='Import',
												  command=self.process)
		self.cmd_process_results.grid(row=self.row(1), column=self.col())

		logger.debug('GUI window created')
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

	def add_new_file(self):
		"""
			Function to load Tkinter.askopenfilename for the user to select a file and then adds to the
			self.file_list scrolling text box
		:return: None
		"""
		# Ask user to select file(s) or folders based on <.bo_files>
		if self.bo_files:
			# User will be able to select list of files
			file_paths = file_selector(initial_pth=self.results_pth, open_file=True)
		else:
			# User will be able to add folder
			file_paths = file_selector(initial_pth=self.results_pth, save_dir=True,
									  lbl_folder_select='Select folder containing hast results files')

		# User can select multiple and so following loop will add each one
		for file_pth in file_paths:
			logger.debug('Results folder {} added as input folder'.format(file_pth))
			self.results_pth = os.path.dirname(file_pth)

			# Add complete file pth to results list
			if not self.results_files_list:
				# If initial list is empty then will need to replace with initial string
				self.lbl_results_files.delete(1.0, tk.END)

			self.results_files_list.append(file_pth)
			self.lbl_results_files.insert(tk.END,
										  '{} - {}\n'.format(len(self.results_files_list), file_pth))

	def process(self):
		"""
			Function sorts the files list to remove any duplicates and then closes GUI window
		:return: None
		"""
		# Sort results into a single list and remove any duplicates
		self.results_files_list = list(set(self.results_files_list))

		# Ask user to select target folder
		target_file = file_selector(initial_pth=self.results_pth, save_file=True,
									lbl_file_select='Please select file for results')[0]

		# Check if user has input an extension and if not then add it
		file, _ = os.path.splitext(target_file)

		self.target_file = file + self.ext

		# Destroy GUI
		self.master.destroy()

class MainGui:
	"""
		Main class to produce the GUI for user interaction
		Allows the user to set up the parameters and define the cases to run the studies
	"""

	def __init__(
			self, title=constants.GuiDefaults.gui_title, start_directory=os.path.dirname(os.path.realpath(__file__))
	):
		"""
		Initialise GUI
		:param title: (optional) - Title to be used for main window
		"""
		# TODO: How to deal with logger handles
		self.logger = logger
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

		# Specific constants that are populated during Main Gui running
		self.study_settings = list()

		# Constants for styles
		# Style for Loading the SAV Button
		# TODO: Review list, most may not be required
		self.style_load_sav = 'LoadSav.TButton'
		self.style_cmd_buttons = 'General.TButton'
		self.style_cmd_run = 'Run.TButton'
		self.style_rating_options = 'TMenubutton'
		self.style_label_general = 'TLabel'
		self.style_label_res = 'Result.TLabel'
		self.style_label_numgens = 'Gens.TLabel'
		self.style_label_mainheading = 'MainHeading.TLabel'
		self.style_label_subheading = 'SubHeading.TLabel'
		self.style_label_subnames = 'SubstationNames.TLabel'
		self.style_label_version_number = 'Version.TLabel'
		self.style_label_notes = 'Notes.TLabel'
		self.style_label_psc_info = 'PSCInfo.TLabel'
		self.style_label_psc_phone = 'PSCPhone.TLabel'
		self.style_radio_buttons = 'TRadiobutton'
		self.style_check_buttons = 'TCheckbutton'

		# Configure styles
		self.configure_styles()

		# Initialisation
		# Add GUI title
		_ = self.add_main_label(row=self.row(1), col=self.col())

		# Add button for selecting Settings file
		self.button_select_settings = self.add_cmd(
			label=constants.GuiDefaults.button_select_settings_label,
			cmd=self.load_settings_file, tooltip='Click to select the file which contains the study settings to be run'
		)

		# Add button for review / editing settings
		self.button_review_settings = self.add_cmd(
			label='Review / Edit Settings',	cmd=self.review_edit_settings,
			tooltip='Click to review / edit the loaded settings', state=tk.DISABLED
		)


		# Add button to run a pre-case check
		self.button_precase_check = self.add_cmd(
			label='Run Pre-case Check',	cmd=self.run_precase_check,
			tooltip='Click to run a pre-case check (optional)', state=tk.DISABLED
		)

		# Add button for pre_case check results
		self.button_precase_results = self.add_cmd(
			label='Review Pre-Case Check Results',	cmd=self.load_results,
			tooltip='Click to review the pre-case check results', state=tk.DISABLED,
			# Row and column numbers declared, aligns with run pre-case check but offset but 1
			row=self.row(), col=self.col()+1
		)

		# Add button to run studies
		self.button_run_studies = self.add_cmd(
			label='Run Studies',	cmd=self.run_studies,
			tooltip='Click to run studies', state=tk.DISABLED
		)

		# Add button for pre_case check results
		self.button_study_results = self.add_cmd(
			label='Review Study Results',	cmd=self.load_results,
			tooltip='Click to review the study results', state=tk.DISABLED,
			# Row and column numbers declared, aligns with run pre-case check but offset but 1
			row=self.row(), col=self.col()+1
		)

		# Will be populated with buttons that need enabling once the study settings have been imported
		self.buttons_to_enable_level1 = (
			self.button_review_settings, self.button_precase_check, self.button_run_studies
		)

		self.logger.debug('GUI window created')

		# to make sure GUI open infront of PSSE gui
		self.master.deiconify()
		# Produce GUI window
		self.master.mainloop()


	def configure_styles(self):
		"""
			Function configures all the ttk styles used within the GUI
			Further details here:  https://anzeljg.github.io/rin2/book2/2405/docs/tkinter/ttk-style-layer.html
		:return:
		"""
		# TODO: Add more detailed commentary
		# Configure the same font in all labels
		standard_font = constants.GuiDefaults.font_family
		bg_color = constants.GuiDefaults.color_main_window
		_ = ttk.Style().configure('.', font=(standard_font, '8'), background=bg_color)

		# Style for Loading the SAV Button
		_ = ttk.Style().configure(self.style_load_sav, height=2, width=30, color=bg_color)

		# Style for all other buttons
		_ = ttk.Style().configure(self.style_load_sav, height=2, color=bg_color)

		# Style for all other buttons
		_ = ttk.Style().configure(self.style_cmd_run, height=2, width=50, color=bg_color)

		_ = ttk.Style().configure(self.style_rating_options, height=2, width=50, background=bg_color)

		_ = ttk.Style().configure(self.style_label_general, background=bg_color)
		_ = ttk.Style().configure(self.style_label_mainheading, font=(standard_font, '10', 'bold'), background=bg_color)
		_ = ttk.Style().configure(self.style_label_subheading, font=(standard_font, '9'), background=bg_color)
		_ = ttk.Style().configure(self.style_label_version_number, font=(standard_font, '7'), background=bg_color)
		_ = ttk.Style().configure(self.style_label_notes, font=(standard_font, '7'), background=bg_color)

		_ = ttk.Style().configure(
			self.style_label_psc_info, font=constants.GuiDefaults.psc_font,
			color=constants.GuiDefaults.psc_color_web_blue, justify='center', background=bg_color
		)

		_ = ttk.Style().configure(
			self.style_label_psc_phone, font=constants.GuiDefaults.psc_font,
			color=constants.GuiDefaults.psc_color_grey, justify='center', background=bg_color
		)

		_ = ttk.Style().configure(self.style_radio_buttons, background=bg_color)

		_ = ttk.Style().configure(self.style_check_buttons, background=bg_color)

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

	def add_main_label(self, row, col, label=constants.GuiDefaults.gui_title):
		"""
			Function to add the name to the GUI
		:param row: Row number to use
		:param col: Column number to use
		:param str label: (optional) = Label to use for header
		:return ttk.Label lbl:  Reference to the newly created label
		"""
		# Add label with the name to the GUI
		lbl = ttk.Label(self.master, text=label, style=self.style_label_mainheading)
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
			self.master, text=label, command=cmd, style=self.style_load_sav, state=state)
		button.grid(row=row, column=col, padx=5)
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
		# Warning message until function fully implemented
		frame = inspect.currentframe()
		self.logger.warning('Function <{}> not yet implemented'.format(inspect.getframeinfo(frame).function))

		# TODO: Function for pre-case check results review needs to be updates with file path of results

		# Needs to enable the precase check button
		self.button_precase_results.configure(state=tk.NORMAL)

	def load_results(self, results_file):
		"""
			Loads a spreadsheet with the pre-case check results
		:return:
		"""
		# Warning message until function fully implemented
		frame = inspect.currentframe()
		self.logger.warning('Function <{}> not yet implemented'.format(inspect.getframeinfo(frame).function))

	def run_studies(self):
		"""
			Runs the full studies
		:return:
		"""
		# Warning message until function fully implemented
		frame = inspect.currentframe()
		self.logger.warning('Function <{}> not yet implemented'.format(inspect.getframeinfo(frame).function))

		# TODO: Function for post study results needs to be updates with file path of results
		# Needs to enable the precase check button
		self.button_study_results.configure(state=tk.NORMAL)

	def load_settings_file(self):
		"""
			Function to allow the user to select the settings file which then once imported enables further buttons
			The function to run the settings file is housed elsewhere
		:return None:
		"""
		# Minimise main window until settings file is loaded
		self.master.iconify()

		self.button_select_settings.configure(constants.GuiDefaults.button_select_settings_label)

		# Ask user to select file(s) or folders
		pth_settings = tk.filedialog.askopenfilename(
			initialdir=self.init_dir,
			filetypes=constants.GuiDefaults.xlsx_types,
			title="Select the SAV case to be loaded for the studies"
		)

		if not os.path.isfile(pth_settings):
			# If the file is empty or not a genuine file then log a message and return
			self.logger.warning('File {} not found, please select a different file'.format(pth_settings))
		else:
			# Import the settings file and check if import successful
			file_inputs = file_io.Excel()
			# TODO: Need to configure settings to return collection of classes that are better to reference / pandas DataFrames
			self.study_settings = file_inputs.import_excel_harmonic_inputs(pth_workbook=pth_settings)

			if not file_inputs.import_success:
				# If there is an error when importing workbook
				# TODO: UNITTEST - To be created for testing setting import failure
				self.logger.warning(
					(
						'Error when trying to import workbook: {}, see messages above and either select a different '
						'workbook or correct this one'
					).format(pth_settings)
				)

			else:
				# If importing workbook was a success then change the state of the other buttons and coninues
				for button in self.buttons_to_enable_level1:
					button.configure(state=tk.NORMAL)

				# Change label for button to reference the study settings
				file_name = os.path.basename(pth_settings)
				self.button_select_settings.configure(label='Settings file: {} loaded'.format(file_name))

		# Return the parent window
		self.master.deiconify()

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
			# Close window
			self.master.destroy()
			self.abort = True
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