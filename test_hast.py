import hast
import hast.constants as constants
import unittest
import logging

#   ------------------------- UNIT TESTS -------------------------------------------
class TestInitSetup(unittest.TestCase):
	"""
		UnitTest to confirm that init setup works correctly, such as log file and powerfactory access
	"""
	def test_logging_setup(self):
		""" Test initialisation of logger"""
		logger = hast.logger
		logger.info(' -- UNIT TEST --')
		self.assertTrue(type(logger) is logging.RootLogger)
		logger.info(' -- UNIT TEST completed --')
		hast.finish_logging()

	def test_powerfactory_setup(self):
		""" Test powerfactory import works successfully and that it is the correct verson that has been imported """
		powerfactory = hast.setup_powerfactory()
		self.assertTrue(powerfactory.__version__, constants.PowerFactory.version)

if __name__ == '__main__':
	unittest.main()