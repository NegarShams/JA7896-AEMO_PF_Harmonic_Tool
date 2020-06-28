"""
#######################################################################################################################
###													logger.py														###
###		Script deals with writing data to PowerFactory and any processing that takes place which requires 			###
###		interacting with power factory																				###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###																													###
#######################################################################################################################
"""
import logging
import logging.handlers
import sys
import os
import traceback

import pscharmonics.constants as constants

class Logger:
	""" Contained within a class since logger will need to print to both power factory and
		to the various log files
	"""
	logger = None # type: logging.Logger

	def __init__(self, pth_debug_log=str(), pth_progress_log=str(), pth_error_log=str(), app=None):
		"""
			Initialise logger
		:param str pth_progress_log:  Full path to the location of the log file that contains the process messages
		:param str pth_error_log:  Full path to the location of the log file that contains the error messages
		:param str pth_debug_log:  Full path to the location of the log file that contains the error messages
		:param bool debug:  True / False on whether running in debug mode or not
		:param powerfactory app: (optional) - If not None then will use this to provide updates to powerfactory
		"""
		# Attributes used during setup_logging
		self.handler_progress_log = None
		self.handler_debug_log = None
		self.handler_error_log = None
		self.handler_stream_log = None

		# Counter for each error message that occurs
		self.warning_count = 0
		self.error_count = 0
		self.critical_count = 0

		self.pth_debug_log = pth_debug_log
		self.pth_progress_log = pth_progress_log
		self.pth_error_log = pth_error_log
		self.app = app
		self.pf_executed = False
		self.debug_mode = constants.DEBUG

		self.file_handlers=[]

		# Set up logger and establish handle for logger
		self.setup_logging()
		self.initial_log_messages()

	def check_paths_exist(self):
		"""
			Function confirms that the desired paths for the log messages already exist and if not will
			create them / revert to the default locations
		:return None:
		"""
		default_folder = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'logs'))

		# Checks the number of log messages in the default folder and provides appropriate warnings to the user before
		# then deleting the oldest

		# Check a path exists if one has been provided
		if self.pth_progress_log:
			pth, progress_name = os.path.split(self.pth_progress_log)

			if not os.path.isdir(pth):
				self.logger.error(
					(
						'The desired path for the log messages: {} does not exist, instead the log messages will be '
						'saved in the default folder {}'
					).format(pth, default_folder)
				)
				debug_name = os.path.basename(self.pth_debug_log)
				error_name = os.path.basename(self.pth_error_log)

				self.pth_debug_log = os.path.join(default_folder, debug_name)
				self.pth_progress_log = os.path.join(default_folder, progress_name)
				self.pth_error_log = os.path.join(default_folder, error_name)
		else:
			# Check default folder exists and if not create it
			if not os.path.isdir(default_folder):
				os.mkdir(default_folder)
				self.logger.debug('Path for log messages created: {}'.format(default_folder))

			c = constants.General
			self.pth_debug_log = os.path.join(default_folder, '{}_{}.log'.format(c.debug_log, constants.uid))
			self.pth_progress_log = os.path.join(default_folder, '{}_{}.log'.format(c.progress_log, constants.uid))
			self.pth_error_log = os.path.join(default_folder, '{}_{}.log'.format(c.error_log, constants.uid))

		return None

	def pf_terminal_running(self):
		"""
			Function determines whether powerfactory is being run from Python or from PowerFactory.  If it is being
			run from PowerFactory then returns True if run from Python then returns False
		:return bool status:  True = PowerFactory, False = Python terminal
		"""
		# Determines whether Python is running or it is being run from PowerFactory directly, if the former then want
		# to ensure log messages are not sent to StdOut as well as the log store
		if self.app:
			# Returns the currently set interface version or 0 if PowerFactory is started from external and
			interface = self.app.GetInterfaceVersion()
			if interface > 0:
				status = True
			else:
				status = False
		else:
			status = False

		return status

	def setup_logging(self):
		"""
			Function to setup the logging functionality
		:return object logger:  Handle to the logger for writing messages
		"""
		self.logger = logging.getLogger(constants.logger_name)
		self.logger.handlers = []

		# Ensures that even debug messages are captured even if they are not written to log file
		self.logger.setLevel(logging.DEBUG)

		# Produce formatter for log entries
		log_formatter = logging.Formatter(fmt='%(asctime)s - %(levelname)s - %(message)s',
										  datefmt='%Y-%m-%d %H:%M:%S')

		# Confirm log paths exists
		self.check_paths_exist()

		# Add handlers
		self.handler_progress_log = self.get_file_handlers(pth=self.pth_progress_log,
														  min_level=logging.INFO,
														  buffer=True, flush_level = logging.ERROR,
														  formatter=log_formatter)
		self.handler_debug_log = self.get_file_handlers(pth=self.pth_debug_log,
														min_level=logging.DEBUG,
														buffer=True, flush_level=logging.CRITICAL,
														buffer_cap=100000,
														formatter=log_formatter)
		self.handler_error_log = self.get_file_handlers(pth=self.pth_error_log,
														min_level=logging.ERROR,
														formatter=log_formatter)

		#self.handler_stream_log = logging.StreamHandler(sys.stdout)
		self.handler_stream_log = logging.StreamHandler(stream=None)

		# If running in DEBUG mode then will export all the debug logs to the window as well
		self.handler_stream_log.setFormatter(log_formatter)
		if self.debug_mode:
			self.handler_stream_log.setLevel(logging.DEBUG)
		else:
			self.handler_stream_log.setLevel(logging.INFO)

		# Add handlers to logger
		self.logger.addHandler(self.handler_progress_log)
		self.logger.addHandler(self.handler_debug_log)
		self.logger.addHandler(self.handler_error_log)
		self.logger.addHandler(self.handler_stream_log)

		return None

	def initial_log_messages(self):
		"""
			Display initial messages for logger including paths where log files will be stored
		:return:
		"""
		# Initial announcement of directories for log messages to be saved in
		self.info('Path for debug log is {} and will be created if any WARNING messages occur'
				  .format(self.pth_debug_log))
		self.info('Path for process log is {} and will contain all INFO and higher messages'
				  .format(self.pth_progress_log))
		self.info('Path for error log is {} and will be created if any ERROR messages occur'
				  .format(self.pth_error_log))
		self.debug(('Stream output is going to stdout which will only be displayed if DEBUG MODE is True and currently '
				   'it is {}'.format(self.debug_mode)))

		# Ensure initial log messages are created and saved to log file
		self.handler_progress_log.flush()
		return None

	def close_logging(self):
		"""Function closes logging but first removes the debug_handler so that the output is not flushed on
			completion.
		"""
		# Close the debug handler so that no debug outputs will be written to the log files again
		# This is a safe close of the logger and any other close, i.e. an exception will result in writing the
		# debug file.
		# Flush existing progress and error logs
		self.handler_progress_log.flush()
		self.handler_error_log.flush()

		# Specifically remove the debug_handler
		self.logger.removeHandler(self.handler_debug_log)

		# Close and delete file handlers so no more logs will be written to file
		for handler in reversed(self.file_handlers):
			handler.close()
			del handler

	def get_file_handlers(self, pth, min_level=logging.INFO, buffer=False, flush_level=logging.INFO, buffer_cap=10,
						  formatter=logging.Formatter()):
		"""
			Function to a handler to write to the target file with our without a buffer if required
			Files are overwritten if they already exist
		:param str pth:  Path to the file handler to be used
		:param int min_level: (optional=logging.INFO) - Is the minimum level that the file handler should include
		:param bool buffer: (optional=False)
		:param int flush_level: (optional=logging.INFO) - The level at which the log messages should be flushed
		:param int buffer_cap:  (optional=10) - Level at which the buffer empties
		:param logging.Formatter formatter:  (optional=logging.Formatter()) - Formatter to use for the log file entries
		:return: logging.handler handler:  Handle for new logging handler that has been created
		"""
		# Handler for process_log, overwrites existing files and buffers unless error message received
		# delay=True prevents the file being created until a write event occurs


		# TODO: Check if error with logger, could be due to use of append command
		handler = logging.FileHandler(filename=pth, mode='a', delay=True)
		self.file_handlers.append(handler)

		# Add formatter to log handler
		handler.setFormatter(formatter)

		# If a buffer is required then create a new memory handler to buffer before printing to file
		if buffer:
			handler = logging.handlers.MemoryHandler(capacity=buffer_cap,
													 flushLevel=flush_level,
													 target=handler)

		# Set the minimum level that this logger will process things for
		handler.setLevel(min_level)

		return handler

	def debug(self, msg):
		""" Handler for debug messages """
		# Only print output to powerfactory if it has been passed to logger
		if self.app and self.pf_executed and self.debug_mode:
			self.app.PrintPlain(msg)
		self.logger.debug(msg)

	def info(self, msg):
		""" Handler for info messages """
		# Only print output to powerfactory if it has been passed to logger
		if self.app and self.pf_executed:
			self.app.PrintPlain(msg)
		self.logger.info(msg)

	def warning(self, msg):
		""" Handler for warning messages """
		self.warning_count += 1
		if self.app and self.pf_executed:
			self.app.PrintWarn(msg)
		self.logger.warning(msg)

	def error(self, msg):
		""" Handler for warning messages """
		self.error_count += 1
		if self.app and self.pf_executed:
			self.app.PrintError(msg)
		self.logger.error(msg)

	def critical(self, msg):
		""" Critical error has occurred """
		# Get calling function to include in log message
		# noinspection PyProtectedMember
		caller = sys._getframe().f_back.f_code.co_name
		self.critical_count += 1

		if self.app and self.pf_executed:
			try:
				# Try statement since possible that an error has occurred and it might not run
				self.app.PrintError(msg)
			# If attribute doesn't exist then continue
			except AttributeError:
				pass
		self.logger.critical('function <{}> reported {}'.format(caller, msg))

	def exception_handler(self, exception_type, value, tb):
		"""
			If an unhandled exception occurs during running of the code it is directed to here
		:param exception_type:
		:param value: str
		:param tb:
		:return:
		"""
		msg = (
			(
				'Script failed with the following exception raised by Python: {}\n\n'
				'Below is the complete traceback that was created by Python: \n{}'
			).format(value, ''.join(traceback.format_exception(exception_type, value, tb)))
		)


		if self.app and self.pf_executed:
			try:
				# Try statement since possible that an error has occurred and it might not run
				self.app.PrintError(msg)
			# If attribute doesn't exist then continue
			except AttributeError:
				pass

		self.logger.critical(msg=msg)

	def flush(self):
		""" Flush all loggers to file before continuing """
		self.handler_progress_log.flush()
		self.handler_error_log.flush()

	def logging_final_report_and_closure(self):
		"""
			Function reports number of error messages raised and closes down logging
		:return None:
		"""
		if sum([self.warning_count, self.error_count, self.critical_count]) > 1:
			self.logger.info(('Log file closing, there were the following number of important messages: \n'
							  '\t - {} Warning Messages that may be of concern\n'
							  '\t - {} Error Messages that may have stopped the results being produced\n'
							  '\t - {} Critical Messages')
							 .format(self.warning_count, self.error_count, self.critical_count))
		else:
			self.logger.info('Log file closing, there were 0 important messages')
		self.logger.debug('Logging stopped')
		logging.shutdown()

	# def __del__(self):
	# 	"""
	# 		To correctly handle deleting and therefore shutting down of logging module
	# 	:return None:
	# 	"""
	# 	self.logging_final_report_and_closure()

	def __exit__(self):
		"""
			To correctly handle deleting and therefore shutting down of logging module
		:return None:
		"""
		self.logging_final_report_and_closure()