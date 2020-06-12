import unittest
import os

from tests.context import pscharmonics


TESTS_DIR = os.path.join(os.path.dirname(__file__), 'test_files')


# ----- UNIT TESTS -----
class TestPowerFactoryConstants(unittest.TestCase):
	""" Tests that the correct python version can be found """

	def test_confirm_paths_empty_initially(self):
		"""
			Test confirm that initially the paths are empty and then become populate on initialising the
			constants
		"""
		pf_constants = pscharmonics.constants.PowerFactory

		self.assertFalse(pf_constants.dig_path)
		self.assertFalse(pf_constants.dig_python_path)

		pf_constants = pf_constants()

		self.assertTrue(pf_constants.dig_path)
		self.assertTrue(pf_constants.dig_python_path)

	def test_pf_2019_success(self):
		""" Test confirms that powerfactory version 2019 can be found successfully """

		pf_constants = pscharmonics.constants.PowerFactory(year='2019', service_pack='')

		self.assertEqual(pf_constants.target_power_factory, 'PowerFactory 2019')

	def test_pf_version_difference(self):
		""" Test confirms that powerfactory version 2019 cannot be found and so loads PowerFactory 2019 """
		# TODO: Poor test since if newer version installed will still create an error

		pf_constants = pscharmonics.constants.PowerFactory(year='2015', service_pack='5')
		self.assertEqual(pf_constants.target_power_factory, 'PowerFactory 2019')

	def test_pf_2019_python_version_fail(self):
		""" Test confirms that if script run from a non-compatible Python version then an exception is thrown """

		self.assertRaises(EnvironmentError, pscharmonics.constants.PowerFactory, mock_python_version='3.1')


class TestUserGuideExists(unittest.TestCase):
	""" Function confirms that the references user guide actually exists in case it gets deleted"""

	def test_existence(self):
		""" Confirm file exists where it's supposed to """
		user_guide = pscharmonics.constants.General.user_guide_pth

		self.assertTrue(os.path.isfile(user_guide), msg='User guide {} does not exist'.format(user_guide))
