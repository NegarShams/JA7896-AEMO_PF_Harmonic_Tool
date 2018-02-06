"""
	Initialisation for power factory to allow process to be run in parallel
"""

import os
import sys
import time
import constants
import traceback
import unittest
import logging

start_time = (time.strftime("%y_%m_%d_%H_%M_%S"))

def setup_logging(pth_logfile=''):
	"""
		Function to setup the logging functionality
	:param str pth_logfile: (optional='') Full path to logfile if not to use default
	:return: logging.logger logger: Returns handle for new logger just created
	"""
	# TODO:  Currently this file will continue to grow, need something to capture that so that growth can be limited and
	# TODO:  it will overwrite unless there is a critical event
	# TODO: Need to incorporate this logging functionality in with the general HAST logging functionality
	# If no name for log file is provided then use the name of the running script plus the time it was started
	# (assuming that this script was the first one initialised)
	if pth_logfile == '':
		script_name, _ = os.path.splitext(sys.argv[0])
		pth_logfile = '{}{}.log'.format(script_name, start_time)

	logging.basicConfig(filename=pth_logfile,
						level=logging.DEBUG,
						format='%(asctime)s: %(levelname)s - %(message)s')
	# Setup logging handler suitably for export to screen
	formatter = logging.Formatter('%(levelname)s - %(message)s')
	console = logging.StreamHandler()
	console.setFormatter(formatter)

	# Only values above this level will be export to screen
	console.setLevel(logging.INFO)

	# Add handler to default logger
	logging.getLogger('').addHandler(console)

	# Log started message
	logging.info('--\tLogging Started\t--')
	logger = logging.getLogger()
	return logger

def setup_powerfactory():
	"""
		Function deals with setting the correct directories required to run PowerFactory and if it is not possible
		then raises suitable exception to alert user.
	:return powerfactory powerfactory: Returns handle to the imported module powerfactory
	"""
	# TODO: If unable to access folder then should search for PowerFactory folder before returning error
	paths = sys.path
	pf_path = constants.PowerFactory.pf_path
	# Check if power factory path is already in system path before adding to avoid excessive length of system path
	if pf_path not in paths:
		sys.path.append(constants.PowerFactory.dig_path)
	if pf_path not in os.environ['PATH']:
		os.environ['PATH'] = os.environ['PATH'] + ';' + pf_path

	try:
		import powerfactory
	except ImportError:
		logging.critical('Unable to import powerfactory')
		traceback.print_exc()
		raise ImportError(' Could not import powerfactory ')













#   ------------------------- UNIT TESTS -------------------------------------------
class InitSetup(unittest.TestCase):
	"""
		UnitTest to confirm that init setup works correctly, such as log file and powerfactory access
	"""
	def test_logging_setup(self):
		""" Simple test of convex_hull calculation to check it is performing correctly """
		logger = setup_logging()
		logger.info(' -- UNIT TEST --')
		self.assertTrue(type(logger) is logging.Logger)
		logging.info(' -- UNIT TEST completed --')