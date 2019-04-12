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
		self.error_count = 0

		self.pth_debug_log = pth_debug_log
		self.pth_progress_log = pth_progress_log
		self.pth_error_log = pth_error_log
		self.app = app
		self.debug_mode = debug

		self.file_handlers=[]

		# Set up logger a
		self.logger = self.setup_logging()

	def setup_logging(self):
		"""
			Function to setup the logging functionality
		:return object logger:  Handle to the logger for writing messages
		"""
		logger = logging.getLogger('HAST')
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

		self.handler_stream_log = logging.StreamHandler(sys.stdout)

		# If running in DEBUG mode then will export all the debug logs to the window as well
		if self.debug_mode:
			self.handler_stream_log.setLevel(logging.DEBUG)
		else:
			self.handler_stream_log.setLevel(logging.INFO)
		self.handler_stream_log.setFormatter(log_formatter)




		# Add handlers to logger
		logger.addHandler(self.handler_progress_log)
		logger.addHandler(self.handler_debug_log)
		logger.addHandler(self.handler_error_log)
		logger.addHandler(self.handler_stream_log)

		logger.info('Path for debug log is {}'.format(self.pth_debug_log))
		logger.info('Path for process log is {}'.format(self.pth_progress_log))
		logger.info('Path for error log is {}'.format(self.pth_error_log))
		logger.debug('Stream output is going to stdout')
		self.handler_progress_log.flush()

		return logger

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

		# Close and delete file handlers so no more logs will be written to file
		for handler in reversed(self.file_handlers):
			handler.close()
			del handler

		# Specifically remove the debug_handler
		self.logger.removeHandler(self.handler_debug_log)

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
		if self.app:
			self.app.PrintPlain(msg)
		self.logger.info(msg)

	def warning(self, msg):
		""" Handler for warning messages """
		if self.app:
			self.app.PrintWarn(msg)
		self.logger.warning(msg)

	def error(self, msg):
		""" Handler for warning messages """
		self.error_count += 1
		if self.app:
			self.app.PrintError(msg)
		self.logger.error(msg)

	def critical(self, msg):
		""" Critical error has occured """
		# Get calling function to include in log message
		caller = sys._getframe().f_back.f_code.co_name

		if self.app:
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
