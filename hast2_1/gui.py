"""
	Script to handle production of GUI and returning files to be processed together to user
"""
import tkinter as tk
import tkinter.filedialog
import tkinter.scrolledtext
import sys
import os
import logging
import glob
import hast2_1.constants as constants

logger = logging.getLogger()

def file_selector(initial_pth='', open_file=False, save_dir=False,
				  save_file=False,
				  lbl_file_select='Select results file(s) to add',
				  lbl_folder_select='Select folder to store results in',
				  def_ext=constants.ResultsExtract.extension):
	"""
		Function to allow the user to select a file to either open or save
	:param str initial_pth: (optional='') Path to use as starting location
	:param bool open_file: (optional=True) will ask for a file to open
	:param bool save_dir: (optional=False) will ask for a directory into which the results should be saved
	:param bool save_file: (optional=False) will ask for a file name to save the results as
	:param str lbl_file_select: (optional) = Title for dialog box that pops up and asks for user to select results file
	:param str lbl_folder_select: (optional) = Title for dialog box when asking user to select a folder
	:param str def_ext: (optional) = Default extension for results export
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
											   filetypes=(('CSV files', '*.csv'),
														  ('All Files', '*.*')))
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
				 def_ext=constants.ResultsExtract.extension):
		"""
			Initialise GUI
		:param str title: (optional) - Title to be used for main window
		:param str start_directory: (optional) - Path to a directory to use for processing
		:param bool files: (optional) = False, when set to False will ask user to select Folders rather than files
		:param str def_ext: (optiona) = '.xlsx', extension of files to search for when using a folder selection
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
