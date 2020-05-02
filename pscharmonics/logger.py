"""
#######################################################################################################################
###											logger																	###
###		Script deals with the logging of messages to both a log handler and to PowerFactory if powerfactory.py is 	###
###		active																										###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###		project JI6973 for EirGrid project PSPF010 - Specialise Support in Power Quality Analysis during 2018		###
###																													###
#######################################################################################################################

"""
import logging
import logging.handlers
import pscharmonics.constants as constants
import sys
import os
import unittest

# Meta Data
__author__ = 'David Mills'
__version__ = '1.3a'
__email__ = 'david.mills@pscconsulting.com'
__phone__ = '+44 7899 984158'
__status__ = 'In Development - Alpha'

class Logger:
	""" Contained within a class since logger will need to print to both power factory and
		to the various log files
	"""
	def __init__(self, pth_debug_log, pth_progress_log, pth_error_log, app=None, debug=False):
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
		self.debug_mode = debug

		self.file_handlers=[]

		# Determine status of whether PowerFactory is running script or if being run from Python
		self.pf_executed = self.pf_terminal_running()

		# Set up logger and establish handle for logger
		self.logger = self.setup_logging()
		self.initial_log_messages()

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
			# SetInerfaceVersion() is not called
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
		# logging.getLogger().setLevel(logging.CRITICAL)
		# logging.getLogger().disabled = True
		logger = logging.getLogger(constants.logger_name)
		logger.handlers = []
		# Ensures that even debug messages are captured even if they are not written to log file

		logger.setLevel(logging.DEBUG)

		# Produce formatter for log entries
		log_formatter = logging.Formatter(fmt='%(asctime)s - %(levelname)s - %(message)s',
										  datefmt='%Y-%m-%d %H:%M:%S')

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
		logger.addHandler(self.handler_progress_log)
		logger.addHandler(self.handler_debug_log)
		logger.addHandler(self.handler_error_log)
		logger.addHandler(self.handler_stream_log)

		return logger

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
		# Debug messages only written to logger
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
		""" Critical error has occured """
		# Get calling function to include in log message
		caller = sys._getframe().f_back.f_code.co_name
		self.critical_count += 1

		if self.app and self.pf_executed:
			try:
				# Try statement since possible that an error has occured and it might not run
				self.app.PrintError(msg)
			# If attribute doesn't exist then continue
			except AttributeError:
				pass
		self.logger.critical('function <{}> reported {}'.format(caller, msg))

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

	def __del__(self):
		"""
			To correctly handle deleting and therefore shutting down of logging module
		:return None:
		"""
		self.logging_final_report_and_closure()

	def __exit__(self):
		"""
			To correctly handle deleting and therefore shutting down of logging module
		:return None:
		"""
		self.logging_final_report_and_closure()

#  ----- UNIT TESTS -----
class TestLoggerSetup(unittest.TestCase):
	"""
		UnitTest to test the operation of various excel workbook functions
	"""
	def setUp(self):
		logging.shutdown()
		# Get path to file that is running
		pth = os.path.dirname(os.path.abspath(__file__))
		pth_test_file = 'test_file_store'
		self.debug_log = os.path.join(pth, pth_test_file, 'debug_log.log')
		self.progress_log = os.path.join(pth, pth_test_file, 'process_log.log')
		self.error_log = os.path.join(pth, pth_test_file, 'error_log.log')

		for file in (self.progress_log, self.debug_log, self.error_log):
			if os.path.isfile(file):
				os.remove(file)

		self.logger = Logger(pth_debug_log=self.debug_log,
							 pth_progress_log=self.progress_log,
							 pth_error_log=self.error_log,
							 app=None)

	def test_progress_logger(self):
		"""
			Tests that only a process log file is produced
		"""
		self.logger.debug('Test debug message')
		self.logger.info('Test info message')

		self.logger.close_logging()
		self.assertTrue(os.path.isfile(self.progress_log))
		self.assertFalse(os.path.isfile(self.debug_log))
		self.assertTrue(self.logger.error_count==0)

	def test_error_logger(self):
		self.logger.debug('Test debug message')
		self.logger.info('Test info message')
		self.logger.warning('Test warning message')
		self.logger.error('Test error message')

		self.logger.close_logging()
		self.assertTrue(os.path.isfile(self.error_log))
		self.assertFalse(os.path.isfile(self.debug_log))
		self.assertTrue(self.logger.error_count == 1)

	def test_debug_logger(self):
		self.logger.debug('Test debug message')
		self.logger.info('Test info message')
		self.logger.warning('Test warning message')
		self.logger.error('Test error message')
		self.logger.critical('Test critical message')

		self.logger.close_logging()
		self.assertTrue(os.path.isfile(self.debug_log))
		self.assertTrue(self.logger.error_count == 1)

	def tearDown(self):
		self.logger.close_logging()
