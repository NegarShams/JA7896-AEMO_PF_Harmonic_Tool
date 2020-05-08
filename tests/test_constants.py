import unittest
import os
import sys
import shutil

from .context import pscharmonics


TESTS_DIR = os.path.join(os.path.dirname(__file__), 'test_files')


# ----- UNIT TESTS -----
class TestPowerFactoryConstants(unittest.TestCase):
	""" Tests that the correct python version can be found """
	def test_pf_2019_success(self):
		""" Test confirms that powerfactory version 2019 can be found successfully """

		pf_constants = pscharmonics.constants.PowerFactory(year='2019', service_pack='')

		self.assertEqual(pf_constants.target_power_factory, 'PowerFactory 2019')

	def test_pf_2019_version_difference(self):
		""" Test confirms that powerfactory version 2019 can be found successfully """

		pf_constants = pscharmonics.constants.PowerFactory(year='2019', service_pack='5')
		self.assertEqual(pf_constants.target_power_factory, 'PowerFactory 2019')

	def test_pf_2019_python_version_fail(self):
		""" Test confirms that powerfactory version 2019 can be found successfully """

		original_minor = sys.version_info.minor

		sys.version_info.minor = 1
		self.assertRaises(EnvironmentError, pscharmonics.constants.PowerFactory)

		# restore version
		sys.version_info.minor = original_minor



