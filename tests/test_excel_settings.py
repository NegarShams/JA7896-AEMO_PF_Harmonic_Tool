"""
	Test module to test importing settings files
"""

import unittest
import os
import pandas as pd

from .context import pscharmonics

TESTS_DIR = os.path.join(os.path.dirname(__file__), 'test_files')

# ----- UNIT TESTS -----
class TestInputs(unittest.TestCase):
	@classmethod
	def setUpClass(cls):
		"""
			Load the SAV case into PSSE for further testing
		"""
		# Initialise logger
		uid = 'Test_Default_Inputs'
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')
		cls.inputs = pscharmonics.file_io.StudyInputsDev(pth_file=pth_inputs)


	def test_study_settings_import(self):
		"""
			unittest to check that still runs if an old HAST Inputs format is used
		:return:
		"""
		# Import the study settings file
		self.inputs.study_settings()

class TestStudySettings(unittest.TestCase):
	"""
		Tests that the class to process all of the study settings works correctly
	"""
	@classmethod
	def setUpClass(cls):
		# Shortening of reference to class and functions under test
		cls.test_cls = pscharmonics.file_io.StudySettings

	def test_initial_fail(self):
		""" Confirms that will not work if a file or workbook is not passed """
		self.assertRaises(IOError, self.test_cls)

	def test_dataframe_import_from_file(self):
		""" Confirm DataFrame imported when loaded using a file """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')

		study_settings = self.test_cls(pth_file=pth_inputs)
		df = study_settings.df

		self.assertEqual(df.shape, (8,1))

	def test_dataframe_import_from_wkbk(self):
		""" Confirm DataFrame imported when loaded using a pandas workbook """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')

		with pd.ExcelFile(pth_inputs) as wkbk:
			study_settings = self.test_cls(wkbk=wkbk)

		df = study_settings.df

		self.assertEqual(df.shape, (8,1))

	def test_dataframe_process_export_folder(self):
		""" Confirm DataFrame imported when loaded using a pandas workbook """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')

		with pd.ExcelFile(pth_inputs) as wkbk:
			study_settings = self.test_cls(wkbk=wkbk)

		# Prove with an empty string
		study_settings.df.loc[study_settings.c.export_folder] = ''
		folder = study_settings.process_export_folder()

		# Expected result is the parent folder of this test file
		expected_folder = os.path.normpath(os.path.join(os.path.dirname(__file__), '..'))
		self.assertEqual(folder, expected_folder)


