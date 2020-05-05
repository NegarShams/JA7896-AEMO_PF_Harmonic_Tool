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

	def test_study_case_no_inputs(self):
		""" Confirm that if no inputs provided then will raise error """
		self.assertRaises(IOError, self.inputs.process_study_cases)

	def test_study_case_duplicated_input(self):
		""" Confirm is processed correctly """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')
		dict_cases = self.inputs.process_study_cases(pth_file=pth_inputs)

		# Confirm length as expected
		self.assertEqual(len(dict_cases), 2)

		# Confirm keys as expected and in correct order
		self.assertEqual(list(dict_cases.keys())[1], 'BASE(1)')

	def test_overall_import(self):
		""" Spot check some of the values are as expected """

		self.assertEqual(self.inputs.cases['BASE'].name, 'BASE')



class GeneralTests(unittest.TestCase):
	"""
		This class is for testing functions which are stand-alone and not part of any
		of the other classes being tested
	"""
	def test_duplicate_entry_updates(self):
		""" Test that duplicate entries in a DataFrame are correctly changed """
		# Create dataset for testing
		key = 'KEY'
		data = [['A', 1, 2],['B',2,3],['A',5,6]]
		columns = ('KEY', 'OTHER', 'OTHER2')

		df_original = pd.DataFrame(data=data, columns=columns)
		df_new, updated = pscharmonics.file_io.update_duplicates(key=key, df=df_original)

		# Confirm no changes in shape after update
		self.assertEqual(df_original.shape, df_new.shape)
		self.assertTrue(updated)
		print(df_new)
		# Confirm a single value
		self.assertTrue('A(1)' in df_new[key].values)

		# Confirm order remains the same
		self.assertTrue(df_original.index.equals(df_new.index))


		# Repeat test with new DataFrame and confirm get no update
		_, updated = pscharmonics.file_io.update_duplicates(key=key, df=df_new)
		self.assertFalse(updated)

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

		self.assertEqual(df.shape, (8, ))

	def test_dataframe_import_from_wkbk(self):
		""" Confirm DataFrame imported when loaded using a pandas workbook """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')

		with pd.ExcelFile(pth_inputs) as wkbk:
			study_settings = self.test_cls(wkbk=wkbk)

		df = study_settings.df

		self.assertEqual(df.shape, (8, ))

	def test_dataframe_process_export_folder(self):
		""" Confirm the file name is correct """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')

		with pd.ExcelFile(pth_inputs) as wkbk:
			study_settings = self.test_cls(wkbk=wkbk)

		# Prove with an empty string
		study_settings.df.loc[study_settings.c.export_folder] = ''
		folder = study_settings.process_export_folder()

		# Expected result is the parent folder of this test file
		expected_folder = os.path.normpath(os.path.join(os.path.dirname(__file__), '..'))
		self.assertEqual(folder, expected_folder)

	def test_dataframe_process_results_name(self):
		""" Confirm the results name is correct """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')
		c = pscharmonics.constants

		with pd.ExcelFile(pth_inputs) as wkbk:
			study_settings = self.test_cls(wkbk=wkbk)

		# Prove with an empty string and fixed UID
		study_settings.df.loc[study_settings.c.results_name] = ''
		study_settings.uid = 'TEST'
		result = study_settings.process_result_name()

		# Expected result is the name below
		expected_result = '{}_TEST.xlsx'.format(c.StudySettings.def_results_name)
		self.assertEqual(result, expected_result)

	def test_blank_network_folder(self):
		""" Confirm that if no network folder is provided raise an Error """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')

		with pd.ExcelFile(pth_inputs) as wkbk:
			study_settings = self.test_cls(wkbk=wkbk)

		study_settings.df.loc[study_settings.c.pf_network_elm] = ''

		self.assertRaises(ValueError, study_settings.process_net_elements)

	def test_get_expected_network_folder(self):
		""" Confirm that if no network folder is provided raise an Error """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')
		expected_result = 'NSW.ElmNet'

		with pd.ExcelFile(pth_inputs) as wkbk:
			study_settings = self.test_cls(wkbk=wkbk)

		# Confirm get the value that is in the Inputs sheet without addition value
		result = study_settings.process_net_elements()
		self.assertEqual(result, expected_result)

		# Remove '.ElmNet' and confirm it gets added back in
		study_settings.df.loc[study_settings.c.pf_network_elm] = 'NSW'

		result = study_settings.process_net_elements()
		self.assertEqual(result, expected_result)

	def test_boolean_sanity_check(self):
		""" Function to confirm that you get correct error messages """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')

		with pd.ExcelFile(pth_inputs) as wkbk:
			study_settings = self.test_cls(wkbk=wkbk)

		# Confirm throws an error if unexpected inputs provided
		study_settings.export_to_excel = False
		study_settings.delete_created_folders = True
		self.assertRaises(ValueError, study_settings.boolean_sanity_check)

		# Confirm that standard response is None
		study_settings.export_to_excel = True
		study_settings.pre_case_check = False
		self.assertIsNone(study_settings.boolean_sanity_check())

	def test_process_booleans(self):
		""" Tests behaviour of Boolean processing function operates correctly """
		""" Function to confirm that you get correct error messages """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')

		with pd.ExcelFile(pth_inputs) as wkbk:
			study_settings = self.test_cls(wkbk=wkbk)

		# Initially expect value to be True
		value = study_settings.process_booleans(key=pscharmonics.constants.StudySettings.export_to_excel)
		self.assertTrue(value)

		# Change value to blank and run again
		study_settings.df.loc[pscharmonics.constants.StudySettings.export_to_excel] = ''
		value = study_settings.process_booleans(key=pscharmonics.constants.StudySettings.export_to_excel)
		self.assertFalse(value)

	def test_single_input_run(self):
		""" Function confirms that all inputs run correctly """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')

		with pd.ExcelFile(pth_inputs) as wkbk:
			study_settings = self.test_cls(wkbk=wkbk)

		# Test just confirms that runs correctly
		self.assertIsNone(study_settings.process_inputs())

