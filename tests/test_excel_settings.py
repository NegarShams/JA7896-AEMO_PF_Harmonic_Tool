"""
	Test module to test importing settings files
"""

import unittest
import os
import pandas as pd
import math
import shutil
import random
import string

from tests.context import pscharmonics

TESTS_DIR = os.path.join(os.path.dirname(__file__), 'test_files')
def_inputs_file = os.path.join(TESTS_DIR, 'Inputs.xlsx')

# Some folders are created during running and these will be deleted
delete_created_folders = True

# noinspection PyUnusedLocal
def mock_process_export_folder(*args, **kwargs):
	"""
		Mock function to deal with processing of the export folder without creating a
		new folder every time the settings are imported
	:param args:
	:param kwargs:
	:return:
	"""

	# Rather than creating a new folder just return reference to the tests directory
	return TESTS_DIR


# ----- MOCKS ------
# Mocks to replace normal functionality
pscharmonics.file_io.StudySettings.process_export_folder = mock_process_export_folder

# ----- UNIT TESTS -----
class TestInputs(unittest.TestCase):
	""" Test the input processing """
	@classmethod
	def setUpClass(cls):
		"""
			Load the SAV case into PSSE for further testing
		"""
		# Initialise logger
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')
		cls.inputs = pscharmonics.file_io.StudyInputs(pth_file=pth_inputs)

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
		self.assertEqual(list(dict_cases.index)[1], 'BASE(1)')

	def test_overall_import(self):
		""" Spot check some of the values are as expected """

		self.assertEqual(self.inputs.cases.loc['BASE', pscharmonics.constants.StudySettings.name], 'BASE')

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
		expected_folder = TESTS_DIR
		self.assertEqual(folder, expected_folder)

	# def test_dataframe_process_results_name(self):
	# 	""" Confirm the results name is correct """
	# 	pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')
	# 	c = pscharmonics.constants
	#
	# 	with pd.ExcelFile(pth_inputs) as wkbk:
	# 		study_settings = self.test_cls(wkbk=wkbk)
	#
	# 	# Prove with an empty string and fixed UID
	# 	study_settings.df.loc[study_settings.c.results_name] = ''
	# 	study_settings.uid = 'TEST'
	# 	result = study_settings.process_result_name()
	#
	# 	# Expected result is the name below
	# 	expected_result = '{}_TEST.xlsx'.format(c.StudySettings.def_results_name)
	# 	self.assertEqual(result, expected_result)

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

class TestContingencies(unittest.TestCase):
	""" Class to deal with testing the reading and processing of contingencies """
	@classmethod
	def setUpClass(cls):
		# Shortening of reference to class and functions under test
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')
		cls.test_cls = pscharmonics.file_io.StudyInputs(pth_inputs)

	def test_dataframe_import_from_file(self):
		""" Confirm DataFrame imported when loaded using a file """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs_cmds.xlsx')

		contingency_cmd, contingencies = self.test_cls.process_contingencies(pth_file=pth_inputs)

		# Confirm command for contingency analysis is imported correctly
		self.assertEqual(contingency_cmd, 'Contingency Analysis')

		# Confirm contingencies are all base case
		self.assertTrue('TEST Cont' in contingencies.keys())

		couplers = contingencies['Base_Case'].couplers
		for coupler in couplers:
			self.assertTrue(math.isnan(coupler.breaker))
			self.assertTrue(math.isnan(coupler.status))

class TestTerminals(unittest.TestCase):
	""" Class to deal with testing the reading and processing of terminals """
	@classmethod
	def setUpClass(cls):
		# Shortening of reference to class and functions under test
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')
		cls.test_cls = pscharmonics.file_io.StudyInputs(pth_inputs)

	def test_dataframe_import_from_file(self):
		""" Confirm DataFrame imported when loaded using a file """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')
		test_term = 'CRANBOURNE 220KV'

		terminals = self.test_cls.process_terminals(pth_file=pth_inputs)

		# Confirm contingencies are all base case
		self.assertTrue(test_term in terminals.keys())

		terminal = terminals[test_term]
		self.assertEqual(terminal.name, test_term)
		self.assertEqual(terminal.substation, '{}.{}'.format(test_term, pscharmonics.constants.PowerFactory.pf_substation))
		self.assertEqual(terminal.terminal, '2958760_CBTS_7MN2.{}'.format(pscharmonics.constants.PowerFactory.pf_terminal))
		self.assertFalse(terminal.include_mutual)

class TestResultsExport(unittest.TestCase):
	""" Testing that the routines associated with exporting the results works correctly """
	def test_results_folder_creation(self):
		""" Confirm that results folder is created correctly """
		# Directory name to be created
		pth = TESTS_DIR
		name = 'test_results'
		target_pth = os.path.join(pth, name)

		# Confirm doesn't already exist
		if os.path.exists(target_pth):
			shutil.rmtree(target_pth)
		self.assertFalse(os.path.exists(target_pth))

		# Initialise class
		res_export = pscharmonics.file_io.ResultsExport(pth=pth, name=name)
		# Create folder
		res_export.create_results_folder()

		# Confirm exists
		self.assertTrue(os.path.isdir(target_pth))

		# Delete folder and confirm deleted
		shutil.rmtree(target_pth)
		self.assertFalse(os.path.exists(target_pth))

class TestLoadFlowSettings(unittest.TestCase):
	""" Testing that load flow settings can be correctly imported with or without a fixed command """

	def setUp(self):
		"""
			Initialisation carried out for every test
		:return:
		"""
		# Following command imports the Load Flow settings from the default inputs spreadsheet
		# TODO: Must ensure reference busbar is defined correctly and default settings match PowerFactory case
		with pd.ExcelFile(def_inputs_file) as wkbk:
			# Import here should match pscconsulting.file_io.StudyInputs().process_lf_settings
			self.df = pd.read_excel(
				wkbk,
				sheet_name=pscharmonics.constants.StudyInputs.lf_settings,
				usecols=(3,), skiprows=3, header=None, squeeze=True
			)

	def test_complete_inputs(self):
		""" Tests the complete set of default inputs """

		# Create instance with complete set of settings
		lf_settings = pscharmonics.file_io.LFSettings(existing_command=self.df.iloc[0], detailed_settings=self.df.iloc[1:])

		# Confirm value of some inputs
		# TODO: Add in a manual check to confirm that all settings are correct
		self.assertFalse(lf_settings.cmd is None)
		self.assertEqual(lf_settings.iopt_net, 0)
		self.assertEqual(lf_settings.iopt_at, 1)

	def test_without_cmd(self):
		""" Tests importing data while missing a pre-populated command """

		# Create instance with complete set of settings
		lf_settings = pscharmonics.file_io.LFSettings(existing_command='', detailed_settings=self.df.iloc[1:])

		# Confirm value of some inputs
		self.assertTrue(lf_settings.cmd is None)
		self.assertEqual(lf_settings.iopt_net, 0)
		self.assertEqual(lf_settings.iopt_at, 1)
		self.assertFalse(lf_settings.settings_error)

	def test_missing_input_with_cmd(self):
		""" Tests importing data while missing a pre-populated command """

		# Create instance with complete set of settings
		lf_settings = pscharmonics.file_io.LFSettings(existing_command=self.df.iloc[0], detailed_settings=self.df.iloc[2:])

		# Confirm value of some inputs
		self.assertFalse(lf_settings.cmd is None)
		self.assertTrue(lf_settings.settings_error)

	def test_missing_input_without_cmd(self):
		""" Tests importing data while missing a pre-populated command without a complete dataset results in
			a incomplete flag and error message the default load flow command will be used
		"""

		# Create instance with complete set of settings
		lf_settings = pscharmonics.file_io.LFSettings(existing_command='', detailed_settings=self.df.iloc[2:])

		# Confirm value of some inputs
		self.assertTrue(lf_settings.cmd is None)
		self.assertTrue(lf_settings.settings_error)

class TestFreqSweepSettings(unittest.TestCase):
	""" Testing that frequency sweep settings can be correctly imported with or without a fixed command """

	def setUp(self):
		"""
			Initialisation carried out for every test
		:return:
		"""
		# Following command imports the Load Flow settings from the default inputs spreadsheet
		# TODO: Must ensure reference busbar is defined correctly and default settings match PowerFactory case
		with pd.ExcelFile(def_inputs_file) as wkbk:
			# Import here should match pscconsulting.file_io.StudyInputs().process_lf_settings
			self.df = pd.read_excel(
				wkbk,
				sheet_name=pscharmonics.constants.StudyInputs.fs_settings,
				usecols=(3,), skiprows=3, header=None, squeeze=True
			)

	def test_complete_inputs(self):
		""" Tests the complete set of default inputs """

		# Create instance with complete set of settings
		settings = pscharmonics.file_io.FSSettings(existing_command=self.df.iloc[0], detailed_settings=self.df.iloc[1:])

		# Confirm value of some inputs
		self.assertFalse(settings.cmd is None)

		# Confirm all following settings match with inputs spreadsheet as test
		self.assertEqual(settings.frnom, 50.0)
		self.assertEqual(settings.iopt_net, 0)
		self.assertEqual(settings.fstart, 50.0)
		self.assertEqual(settings.fstop, 500.0)
		self.assertEqual(settings.fstep, 50.0)
		self.assertEqual(settings.i_adapt, False)
		self.assertEqual(settings.errmax, 0.01)
		self.assertEqual(settings.errinc, 0.005)
		self.assertEqual(settings.ninc, 10.0)
		self.assertEqual(settings.ioutall, False)


	def test_without_cmd(self):
		""" Tests importing data while missing a pre-populated command """

		# Create instance with complete set of settings
		settings = pscharmonics.file_io.FSSettings(existing_command='', detailed_settings=self.df.iloc[1:])

		# Confirm value of some inputs
		self.assertTrue(settings.cmd is None)

	def test_missing_input_with_cmd(self):
		""" Tests importing data while missing a pre-populated command """

		# Create instance with complete set of settings
		settings = pscharmonics.file_io.FSSettings(existing_command=self.df.iloc[0], detailed_settings=self.df.iloc[2:])

		# Confirm value of some inputs
		self.assertFalse(settings.cmd is None)
		self.assertTrue(settings.settings_error)

	def test_missing_input_without_cmd(self):
		""" Tests importing data while missing a pre-populated command without a complete dataset results in
			a incomplete flag and error message the default load flow command will be used
		"""

		# Create instance with complete set of settings
		fs_settings = pscharmonics.file_io.FSSettings(existing_command='', detailed_settings=self.df.iloc[2:])

		# Confirm value of some inputs
		self.assertTrue(fs_settings.cmd is None)
		self.assertTrue(fs_settings.settings_error)

class TestDeleteOldFiles(unittest.TestCase):
	"""
		Tests the function for deleting of files which are greater than a particular number
	"""
	def setUp(self):
		""" Creates a temporary folder for testing """

		self.temp_folder = os.path.join(
			TESTS_DIR,
			'test_folder_{}'.format(''.join(random.choice(string.ascii_uppercase + string.digits) for _ in range(3)))
		)

		# Creates folder if doesn't already exist
		if not os.path.isdir(self.temp_folder):
			os.mkdir(self.temp_folder)

		# Threshold values for testing (warning, error)
		self.thresholds = (20, 30)


	def test_path_does_not_exist(self):
		""" Tests that if path doesn't exist just returns """

		# Delete folder and confirm doesn't exist
		shutil.rmtree(self.temp_folder)
		self.assertFalse(os.path.isdir(self.temp_folder))

		# Test deletion
		num_deleted = pscharmonics.file_io.delete_old_files(pth=self.temp_folder, logger=pscharmonics.constants.logger)

		self.assertTrue(num_deleted==0)

	def test_below_thresholds(self):
		""" Tests that if below threshold doesn't delete files"""

		# Delete folder and confirm doesn't exist
		for x in range(10):
			file_pth = os.path.join(self.temp_folder, 'file{}.txt'.format(x))
			with open(file_pth, 'w') as f:
				f.write('test file')

		# Test deletion
		num_deleted = pscharmonics.file_io.delete_old_files(
			pth=self.temp_folder, logger=pscharmonics.constants.logger, thresholds=self.thresholds)

		self.assertTrue(num_deleted==0)

	def test_warning_threshold(self):
		""" Tests that if below threshold doesn't delete files"""

		# Delete folder and confirm doesn't exist
		for x in range(self.thresholds[0]+1):
			file_pth = os.path.join(self.temp_folder, 'file{}.txt'.format(x))
			with open(file_pth, 'w') as f:
				f.write('test file')

		# Test deletion
		num_deleted = pscharmonics.file_io.delete_old_files(
			pth=self.temp_folder, logger=pscharmonics.constants.logger, thresholds=self.thresholds)

		self.assertTrue(num_deleted==0)

	def test_delete_threshold(self):
		""" Tests that if below threshold doesn't delete files"""

		# Delete folder and confirm doesn't exist
		for x in range(self.thresholds[1]):
			file_pth = os.path.join(self.temp_folder, 'file{}.txt'.format(x))
			with open(file_pth, 'w') as f:
				f.write('test file')

		# Test deletion
		num_deleted = pscharmonics.file_io.delete_old_files(
			pth=self.temp_folder, logger=pscharmonics.constants.logger, thresholds=self.thresholds)

		self.assertTrue(num_deleted==self.thresholds[1]-self.thresholds[0])

	def test_delete_threshold_using_constants(self):
		""" Tests that if below threshold doesn't delete files"""

		# Delete folder and confirm doesn't exist
		for x in range(pscharmonics.constants.General.threshold_delete):
			file_pth = os.path.join(self.temp_folder, 'file{}.txt'.format(x))
			with open(file_pth, 'w') as f:
				f.write('test file')

		# Test deletion
		num_deleted = pscharmonics.file_io.delete_old_files(pth=self.temp_folder, logger=pscharmonics.constants.logger)

		expected_delete = pscharmonics.constants.General.threshold_delete - pscharmonics.constants.General.threshold_warning
		self.assertTrue(num_deleted==expected_delete)


	def tearDown(self):
		""" Deletes the entire folder """
		if os.path.isdir(self.temp_folder):
			shutil.rmtree(self.temp_folder)

class TestCombineResults(unittest.TestCase):
	"""
		Tests the function for deleting of files which are greater than a particular number
	"""
	def setUp(self):
		""" Check previous results already exist """

		self.results1 = os.path.join(TESTS_DIR, 'Detailed_1')
		self.results2 = os.path.join(TESTS_DIR, 'Detailed_2')

		for x in (self.results1, self.results2):
			self.assertTrue(
				os.path.isdir(x),
				msg='The detailed results folder {} does not exist, run <test_pf.py> first to '
					'produce'
			)


	def test_export_single_results_set(self):
		""" Tests exporting of a single results set works """

		src_paths = (self.results1, )
		target_file = os.path.join(TESTS_DIR, 'combined_results1.xlsx')

		# Confirm file doesn't already exist
		if os.path.isfile(target_file):
			os.remove(target_file)

		pscharmonics.file_io.ExtractResults(target_file=target_file, search_paths=src_paths)

	def test_export_single_results_set2(self):
		""" Tests exporting of a single results set works """

		src_paths = (self.results2, )
		target_file = os.path.join(TESTS_DIR, 'combined_results2.xlsx')

		# Confirm file doesn't already exist
		if os.path.isfile(target_file):
			os.remove(target_file)

		pscharmonics.file_io.ExtractResults(target_file=target_file, search_paths=src_paths)

	def test_export_combined_results_set(self):
		""" Tests exporting of a single results set works """

		src_paths = (self.results1, self.results2)
		target_file = os.path.join(TESTS_DIR, 'combined_results3.xlsx')

		# Confirm file doesn't already exist
		if os.path.isfile(target_file):
			os.remove(target_file)

		pscharmonics.file_io.ExtractResults(target_file=target_file, search_paths=src_paths)