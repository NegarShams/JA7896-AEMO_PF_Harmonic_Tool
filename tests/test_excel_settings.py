"""
#######################################################################################################################
###													test_excel_settings.py											###
###		Script deals with testing of importing settings files and processing combined results 						###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###																													###
#######################################################################################################################
"""

import unittest
import os
import pandas as pd
import time
import shutil
import random
import string
import shapely.geometry
import shapely.geometry.polygon
import matplotlib.pyplot
from functools import partial

from tests.context import pscharmonics

TESTS_DIR = os.path.join(os.path.dirname(__file__), 'test_files')
def_inputs_file = os.path.join(TESTS_DIR, 'Inputs.xlsx')

# Some folders are created during running and these will be deleted
delete_created_folders = True

# Set to True if figures should be plotted and displayed
PLOT_FIGURES = True

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

class MockExtractResults:
	""" Mock created to allow independent testing of combine multiple runs """
	def __init__(self):
		self.include_convex = True
		self.combine_multiple_runs = partial(pscharmonics.file_io.ExtractResults.combine_multiple_runs, self)

		self.nom_frequency = float()
		# Create target frequency range
		self.target_freq_range = dict()
		self.percentage_to_exclude = dict()
		self.max_vertices = dict()

class MockPreviousResultsExport:
	""" Mock created to allow independent testing of results processing functions """
	def __init__(self):
		# Constants to be defined
		self.logger = pscharmonics.constants.logger
		self.study_type = 'FS'

		# Functions to be defined
		self.process_file = partial(pscharmonics.file_io.PreviousResultsExport.process_file, self)
		self.process_file_name = partial(pscharmonics.file_io.PreviousResultsExport.process_file_name, self)
		self.extract_var_name = partial(pscharmonics.file_io.PreviousResultsExport.extract_var_name, self)
		self.extract_var_type = partial(pscharmonics.file_io.PreviousResultsExport.extract_var_type, self)


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

		# Updated because inputs now includes a setting for Include Convex
		self.assertEqual(df.shape, (9, ))

	def test_dataframe_import_from_wkbk(self):
		""" Confirm DataFrame imported when loaded using a pandas workbook """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')

		with pd.ExcelFile(pth_inputs) as wkbk:
			study_settings = self.test_cls(wkbk=wkbk)

		df = study_settings.df

		# Updated because inputs now includes a setting for Include Convex
		self.assertEqual(df.shape, (9, ))

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

	def test_include_convex_handled(self):
		""" Function confirms that after update to Include ConvexHull it is now correctly detected in the inputs """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')

		with pd.ExcelFile(pth_inputs) as wkbk:
			study_settings = self.test_cls(wkbk=wkbk)

		self.assertTrue(study_settings.include_convex)

		# Test just confirms that runs correctly
		self.assertIsNone(study_settings.process_inputs())

	def test_include_convex_backward_compatibility(self):
		""" Function confirms that old version of inputs still handled correctly but with convex_hull set to False """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs_v1.0.xlsx')

		with pd.ExcelFile(pth_inputs) as wkbk:
			study_settings = self.test_cls(wkbk=wkbk)

		self.assertFalse(study_settings.include_convex)

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

		# Test one couple exactly matches input requirements
		couplers = contingencies['TEST Cont'].couplers
		coupler = couplers[0]  # type: pscharmonics.file_io.CouplerDetails
		self.assertEqual(coupler.substation, 'ALBURY 132KV')
		self.assertEqual(coupler.breaker, 'Switch_213211')
		self.assertEqual(coupler.status, False)

class TestContingenciesLineData(unittest.TestCase):
	"""
		Class to deal with testing the reading and processing of contingencies that relate to names
		of lines rather than identifying specific circuit breakers
	"""
	@classmethod
	def setUpClass(cls):
		# Shortening of reference to class and functions under test
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs.xlsx')
		cls.test_cls = pscharmonics.file_io.StudyInputs(pth_inputs)

	def test_dataframe_import_from_file(self):
		""" Confirm DataFrame imported when loaded using a file """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs_cont_lines.xlsx')

		contingency_cmd, contingencies = self.test_cls.process_contingencies(pth_file=pth_inputs, line_data=True)

		# Confirm command for contingency analysis is imported correctly
		self.assertEqual(contingency_cmd, 'Contingency Analysis')

		# Confirm contingencies are all base case
		self.assertTrue('TEST Line' in contingencies.keys())

		lines = contingencies['TEST Line'].lines  # type: list
		for line in lines:  # type: pscharmonics.file_io.LineDetails
			self.assertEqual(line.line, '207586_BATS_TGTS_220')
			self.assertEqual(line.status, False)

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

class TestLociSettings(unittest.TestCase):
	""" Class to deal with testing the reading and processing of loci settings """
	def test_import_settings_non_custom(self):
		""" Confirm DataFrame imported when loaded using a file """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs_Detailed6.xlsx')

		loci_settings = pscharmonics.file_io.LociSettings(pth_file=pth_inputs)

		# Confirm command for contingency analysis is imported correctly
		self.assertFalse(loci_settings.custom_polygon)
		self.assertFalse(loci_settings.custom_exclude)

	def test_import_settings_custom(self):
		""" Confirm DataFrame imported when loaded using a file """
		pth_inputs = os.path.join(TESTS_DIR, 'Inputs_loci_custom.xlsx')

		loci_settings = pscharmonics.file_io.LociSettings(pth_file=pth_inputs)

		# Confirm command for contingency analysis is imported correctly
		self.assertTrue(loci_settings.custom_polygon)
		self.assertTrue(loci_settings.custom_exclude)

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
		# Forces a brief wait before attempting to delete file
		time.sleep(0.5)
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

		# Confirm file created
		self.assertTrue(os.path.exists(target_file))

	def test_export_single_results_set2(self):
		""" Tests exporting of a single results set works """

		src_paths = (self.results2, )
		target_file = os.path.join(TESTS_DIR, 'combined_results2.xlsx')

		# Confirm file doesn't already exist
		if os.path.isfile(target_file):
			os.remove(target_file)

		pscharmonics.file_io.ExtractResults(target_file=target_file, search_paths=src_paths)

		# Confirm file created
		self.assertTrue(os.path.exists(target_file))

	def test_export_combined_results_set(self):
		""" Tests exporting of a single results set works """

		src_paths = (self.results1, self.results2)
		target_file = os.path.join(TESTS_DIR, 'combined_results3.xlsx')

		# Confirm file doesn't already exist
		if os.path.isfile(target_file):
			os.remove(target_file)

		pscharmonics.file_io.ExtractResults(target_file=target_file, search_paths=src_paths)

	def test_export_single_results_set3_no_contingencies(self):
		""" Tests exporting of a results set with no contingencies (only the base case) works """

		# Source path to search and confirm exist before continuing
		src_path = os.path.join(TESTS_DIR, 'Detailed_3')
		self.assertTrue(os.path.isdir(src_path))

		# Target file for export
		target_file = os.path.join(TESTS_DIR, 'combined_results_no_cont.xlsx')
		# Confirm file doesn't already exist
		if os.path.isfile(target_file):
			os.remove(target_file)

		pscharmonics.file_io.ExtractResults(target_file=target_file, search_paths=(src_path, ))

		# Confirm file created
		self.assertTrue(os.path.exists(target_file))

class TestResultsProcessing(unittest.TestCase):
	"""
		Tests that results processing of previous raw files works correctly

		Tests carried out on Detailed 1 data processing
	"""
	def setUp(self):
		""" Check previous results already exist """

		test_dir = os.path.join(TESTS_DIR, 'RawResultsFiles')
		test_results1 = os.path.join(test_dir, 'FS_BASE_Intact.csv')
		test_inputs = os.path.join(test_dir, 'InputsDetailed_Results.xlsx')

		for x in (test_results1, ):
			self.assertTrue(
				os.path.isfile(x),
				msg='The previous results export <{}> does not exist, run <test_pf.py> first to produce'.format(test_results1)
			)

		cls_mock = MockPreviousResultsExport()

		cls_mock.inputs = pscharmonics.file_io.StudyInputs(pth_file=test_inputs, gui_mode=True)

		self.df = cls_mock.process_file(pth=test_results1)

	def test_nom_voltage_correct(self):
		""" Tests that the expected nominal voltages are calculated """

		# For node

		pass

		# Extract all nominal voltages and then convert to a unique list
		idx = pd.IndexSlice
		nom_voltages = list()
		for nom_v in self.df.loc[:, idx[:,:,:,:,:,:,pscharmonics.constants.PowerFactory.pf_nom_voltage]].values:
			nom_voltages.extend(nom_v)
		nom_voltages = set(nom_voltages)

		# Confirm number of values returned
		self.assertEqual(len(nom_voltages), 3)
		# Confirm expected values appear in lists
		for value in (11, 220, 330):
			self.assertTrue(
				value in nom_voltages, msg='Expected nominal voltage {} kV not found in returned data frame'.format(value)
			)


class TestCreateConvex(unittest.TestCase):
	""" Tests that passing R/X data will return ConvexHull around data points """

	def setUp(self):
		""" Creates a random data set """
		self.results4 = os.path.join(TESTS_DIR, 'Detailed_4')
		self.results5 = os.path.join(TESTS_DIR, 'Detailed_5')
		self.results5b = os.path.join(TESTS_DIR, 'Detailed_5b')
		self.results6 = os.path.join(TESTS_DIR, 'Detailed_6')

		for x in (self.results4, self.results5, self.results6):
			self.assertTrue(
				os.path.isdir(x),
				msg='The detailed results folder {} does not exist, run <test_pf.py> first to '
					'produce'
			)

		# Detailed 5 and Detailed 5b compared
		self.detailed5_export = os.path.join(TESTS_DIR, 'combined_results_5.xlsx')
		self.detailed5b_export = os.path.join(TESTS_DIR, 'combined_results_5b.xlsx')

		self.cls_extract = MockExtractResults()

	def test_convex_points(self):
		""" Tests can be created with unlimited vertices"""
		# Upper limit of range
		upper_limit = int(pscharmonics.constants.PowerFactory.max_impedance - 1)
		number_points = 50

		x_points = (random.sample(range(upper_limit), number_points))
		y_points = (random.sample(range(upper_limit), number_points))

		corners = pscharmonics.file_io.find_convex_vertices(
			x_values=x_points, y_values=y_points, max_vertices=pscharmonics.constants.LociInputs.unlimited_identifier
		)

		# Confirm all points lie within the Polygon returned by the vertices
		polygon = shapely.geometry.polygon.Polygon(list(zip(*corners)))
		rand_point = random.randint(0, number_points-1)
		point = shapely.geometry.Point(x_points[rand_point], y_points[rand_point])

		self.assertTrue(polygon.contains(point))


		# If required will produce a plot of the required data
		if PLOT_FIGURES:
			matplotlib.pyplot.plot(x_points, y_points, 'o')

			matplotlib.pyplot.plot(corners[0], corners[1], 'r--')
			matplotlib.pyplot.show()

	def test_convex_points_limited_vertices(self):
		""" Tests can be created with maximum of 5 vertices """
		max_vertices = 5
		# Upper limit of range
		upper_limit = int(pscharmonics.constants.PowerFactory.max_impedance - 1)
		number_points = 50

		x_points = (random.sample(range(upper_limit), number_points))
		y_points = (random.sample(range(upper_limit), number_points))

		corners = pscharmonics.file_io.find_convex_vertices(
			x_values=x_points, y_values=y_points, max_vertices=max_vertices
		)

		# Confirm all points lie within the Polygon returned by the vertices
		polygon = shapely.geometry.polygon.Polygon(list(zip(*corners)))
		rand_point = random.randint(0, number_points-1)
		point = shapely.geometry.Point(x_points[rand_point], y_points[rand_point])

		self.assertTrue(polygon.contains(point))

		# If required will produce a plot of the required data
		if PLOT_FIGURES:
			matplotlib.pyplot.plot(x_points, y_points, 'o')

			matplotlib.pyplot.plot(corners[0], corners[1], 'r--')
			matplotlib.pyplot.show()

	def test_convex_points_2_only(self):
		""" Tests can be created with only 2 points"""
		# Upper limit of range
		upper_limit = int(pscharmonics.constants.PowerFactory.max_impedance - 1)
		number_points = 2

		x_points = (random.sample(range(upper_limit), number_points))
		y_points = (random.sample(range(upper_limit), number_points))

		corners = pscharmonics.file_io.find_convex_vertices(
			x_values=x_points, y_values=y_points, max_vertices=pscharmonics.constants.LociInputs.unlimited_identifier
		)

		# Confirm x and y points in list returned
		self.assertTrue(x_points[0] in corners[0])
		self.assertTrue(y_points[0] in corners[1])

	def test_convex_points_0_valid_values(self):
		""" Tests can be created with only 2 points"""
		# Upper limit of range
		corners = pscharmonics.file_io.find_convex_vertices(
			x_values=tuple(), y_values=tuple(), max_vertices=pscharmonics.constants.LociInputs.unlimited_identifier
		)

		# Confirm x and y points in list returned
		self.assertTrue(len(corners[0])==0)
		self.assertTrue(len(corners[1])==0)

	def test_convex_from_data_for_detailed4_data(self):
		"""
			Tests that imported data can be processed and extracted into a suitable DataFrame format
			Note:  A minimum of 3 different data points are needed for a Convex
		"""

		# Import the necessary raw data
		src_paths = (self.results4,)
		df, extract_vars = self.cls_extract.combine_multiple_runs(search_paths=src_paths)

		# Create target frequency range
		target_freq_range = dict()
		percentage_to_exclude = dict()
		max_vertices = dict()
		nom_freq = 50.0
		for h in range(2, 12):
			target_freq_range[h] = (h*nom_freq - nom_freq / 2.0, h*nom_freq + nom_freq/2.0)
			percentage_to_exclude[h] = 0.0
			max_vertices[h] = pscharmonics.constants.LociInputs.unlimited_identifier

		# Pass to function to calculate
		df_convex = pscharmonics.file_io.calculate_convex_vertices(
			df=df, frequency_bounds=target_freq_range, percentage_to_exclude=percentage_to_exclude,
			max_vertices=max_vertices
		)

		# Confirm expected values returned (Expect harmonic order 11 to be empty)
		idx = pd.IndexSlice
		self.assertFalse(df_convex.loc[:, idx['CRANBOURNE 220KV', 'h = 3  (125.0 - 175.0 Hz)', :]].dropna().empty)

	def test_convex_from_data_for_detailed4_data_nonlinear_harmonic_numbers(self):
		"""
			Tests that imported data can be processed to calculate the convex hull for non-linear harmonic numbers
		"""

		# Import the necessary raw data
		src_paths = (self.results4,)
		df, extract_vars = self.cls_extract.combine_multiple_runs(search_paths=src_paths)

		# Create target frequency range
		target_freq_range = dict()
		percentage_to_exclude = dict()
		max_vertices = dict()
		nom_freq = 50.0
		# Produce a dataset for the harmonic numbers which is non-linear
		for h in range(2, 12):
			target_freq_range[h] = (h*nom_freq - nom_freq / float(h), h*nom_freq + nom_freq/float(h))
			percentage_to_exclude[h] = 0.0
			max_vertices[h] = pscharmonics.constants.LociInputs.unlimited_identifier

		# Pass to function to calculate
		df_convex = pscharmonics.file_io.calculate_convex_vertices(
			df=df, frequency_bounds=target_freq_range, percentage_to_exclude=percentage_to_exclude,
			max_vertices=max_vertices
		)

		# Confirm expected values returned (Expect harmonic order 11 to be empty)
		idx = pd.IndexSlice
		self.assertFalse(df_convex.loc[:, idx['CRANBOURNE 220KV', 'h = 2  (75.0 - 125.0 Hz)', :]].dropna().empty)
		self.assertFalse(df_convex.loc[:, idx['CRANBOURNE 220KV', 'h = 4  (187.5 - 212.5 Hz)', :]].dropna().empty)

	def test_convex_from_data_for_detailed4_data_nonlinear_harmonic_numbers_limited_vertices(self):
		"""
			Tests that imported data can be processed to calculate the convex hull for non-linear harmonic numbers
		"""

		# Import the necessary raw data
		src_paths = (self.results4,)
		df, extract_vars = self.cls_extract.combine_multiple_runs(search_paths=src_paths)

		# Create target frequency range
		target_freq_range = dict()
		percentage_to_exclude = dict()
		max_vertices = dict()
		nom_freq = 50.0
		# Produce a dataset for the harmonic numbers which is non-linear
		for h in range(2, 12):
			target_freq_range[h] = (h*nom_freq - nom_freq / float(h), h*nom_freq + nom_freq/float(h))
			percentage_to_exclude[h] = 0.0
			max_vertices[h] = 5

		# Pass to function to calculate
		df_convex = pscharmonics.file_io.calculate_convex_vertices(
			df=df, frequency_bounds=target_freq_range, percentage_to_exclude=percentage_to_exclude,
			max_vertices=max_vertices
		)

		# Confirm expected values returned (Expect harmonic order 11 to be empty)
		idx = pd.IndexSlice
		self.assertFalse(df_convex.loc[:, idx['CRANBOURNE 220KV', 'h = 2  (75.0 - 125.0 Hz)', :]].dropna().empty)
		self.assertFalse(df_convex.loc[:, idx['CRANBOURNE 220KV', 'h = 4  (187.5 - 212.5 Hz)', :]].dropna().empty)

	def test_convex_from_data_for_detailed4_data_percentage_to_exclude(self):
		"""
			Tests that processing of data will exclude certain values from ConvexHull
		"""

		# Import the necessary raw data
		src_paths = (self.results4,)
		df, extract_vars = self.cls_extract.combine_multiple_runs(search_paths=src_paths)

		# Create target frequency range
		target_freq_range = dict()
		percentage_to_exclude = dict()
		max_vertices = dict()
		nom_freq = 50.0
		for h in range(2, 12):
			target_freq_range[h] = (h*nom_freq - nom_freq / 2.0, h*nom_freq + nom_freq/2.0)
			percentage_to_exclude[h] = 0.0
			max_vertices[h] = pscharmonics.constants.LociInputs.unlimited_identifier


		# For the node:  Cranbourne would expect no values greater than 8 for R or X if top 5% excluded for 2nd harmonic
		r_max = 7.9

		idx = pd.IndexSlice
		test_node = 'CRANBOURNE 220KV'
		test_h = 'h = 2  (75.0 - 125.0 Hz)'
		test_r = pscharmonics.constants.PowerFactory.pf_r1
		test_x = pscharmonics.constants.PowerFactory.pf_x1

		# Initially calculate with all data
		df_convex_all = pscharmonics.file_io.calculate_convex_vertices(
			df=df, frequency_bounds=target_freq_range, percentage_to_exclude=percentage_to_exclude,
			max_vertices=max_vertices
		)
		# Obtain maximum values from this DataFrame
		max_values_all = (
			max(df_convex_all.loc[:, idx[test_node, test_h, test_r]]),
			max(df_convex_all.loc[:, idx[test_node, test_h, test_x]]),
		)

		# Confirm that max values exceeds threshold
		self.assertTrue(max_values_all[0]>r_max)

		# Calculate with top 5% excluded
		percentage_to_exclude = {k: 0.05 for k in percentage_to_exclude.keys()}
		df_convex_5 = pscharmonics.file_io.calculate_convex_vertices(
			df=df, frequency_bounds=target_freq_range, percentage_to_exclude=percentage_to_exclude,
			max_vertices=max_vertices
		)
		# Obtain maximum values from this DataFrame
		max_values_5 = (
			max(df_convex_5.loc[:, idx[test_node, test_h, test_r]]),
			max(df_convex_5.loc[:, idx[test_node, test_h, test_x]]),
		)

		# Confirm that max values below threshold
		self.assertTrue(max_values_5[0]<r_max)


		# Calculate with top 50% excluded
		percentage_to_exclude = {k: 0.5 for k in percentage_to_exclude.keys()}
		df_convex_50 = pscharmonics.file_io.calculate_convex_vertices(
			df=df, frequency_bounds=target_freq_range, percentage_to_exclude=percentage_to_exclude,
			max_vertices=max_vertices
		)

		# Obtain maximum values from this DataFrame
		max_values_50 = (
			max(df_convex_50.loc[:, idx[test_node, test_h, test_r]]),
			max(df_convex_50.loc[:, idx[test_node, test_h, test_x]]),
		)

		# Confirm that max values are in correct order
		self.assertTrue(max_values_50[0]<max_values_5[0]<max_values_all[0])
		self.assertTrue(max_values_50[1]<max_values_5[1]<=max_values_all[1])

	def test_raw_r_x_values_for_detailed_4(self):
		"""
			Confirms that the raw_r and raw_x excel references match up for the data provided
		"""

		# Import the necessary raw data
		src_paths = (self.results4,)
		df, extract_vars = self.cls_extract.combine_multiple_runs(search_paths=src_paths)

		# Create target frequency range
		target_freq_range = pscharmonics.file_io.LociSettings().freq_bands

		pscharmonics.file_io.get_raw_data_excel_references(
			sht_name='TEST',
			df=df, start_row=1, start_col=1, target_frequencies=target_freq_range
		)

	def test_export_detailed_results4_including_convex(self):
		""" Tests exporting of a results set with convex hull plots works """

		# Source path to search and confirm exist before continuing
		src_path = (self.results4,)

		# Target file for export
		target_file = os.path.join(TESTS_DIR, 'combined_results_4.xlsx')
		# Confirm file doesn't already exist
		if os.path.isfile(target_file):
			os.remove(target_file)

		# Force to True so results handled correctly
		pscharmonics.file_io.ExtractResults.include_convex = True
		pscharmonics.file_io.ExtractResults(target_file=target_file, search_paths=src_path)

		# Confirm file created
		self.assertTrue(os.path.exists(target_file))

	def test_export_detailed_results5_including_convex(self):
		""" Tests exporting of a results set with convex hull plots works """

		# Source path to search and confirm exist before continuing
		src_path = (self.results5,)

		# Target file for export
		target_file = self.detailed5_export

		# Confirm file doesn't already exist
		if os.path.isfile(target_file):
			os.remove(target_file)

		# Force to True so results handled correctly
		pscharmonics.file_io.ExtractResults.include_convex = True
		pscharmonics.file_io.ExtractResults(target_file=target_file, search_paths=src_path)

		# Confirm file created
		self.assertTrue(os.path.exists(target_file))

	def test_export_detailed_results5b_including_convex(self):
		""" Tests exporting of a results set which should be the same as detailed_5 but results produced
			using a contingency command instead
		"""

		# Source path to search and confirm exist before continuing
		src_path = (self.results5b,)

		# Target file for export
		target_file = self.detailed5b_export
		# Confirm file doesn't already exist
		if os.path.isfile(target_file):
			os.remove(target_file)

		# Force to True so results handled correctly
		pscharmonics.file_io.ExtractResults.include_convex = True
		pscharmonics.file_io.ExtractResults(target_file=target_file, search_paths=src_path)

		# Confirm file created
		self.assertTrue(os.path.exists(target_file))

	def test_export_detailed_results6_including_convex(self):
		""" Tests exporting of a results set with convex hull plots works """

		# Source path to search and confirm exist before continuing
		src_path = (self.results6,)

		# Target file for export
		target_file = os.path.join(TESTS_DIR, 'combined_results_6.xlsx')
		# Confirm file doesn't already exist
		if os.path.isfile(target_file):
			os.remove(target_file)

		# Force to True so results handled correctly
		pscharmonics.file_io.ExtractResults.include_convex = True
		pscharmonics.file_io.ExtractResults(target_file=target_file, search_paths=src_path)

		# Confirm file created
		self.assertTrue(os.path.exists(target_file))

	def test_results_match(self):
		"""
			Test routine to confirm that results produced using either lists of contingencies or the contingencies command
			will exactly match each other.  This test is performed using:
				Detailed5 - Results produced using lists of contingencies
				Detailed5b - Results produced using a pre-determined contingencies command
			NOTE: These must both exist for this test to be carried out
		"""

		# Confirm results already exist and if not run study
		if not os.path.isfile(self.detailed5_export):
			self.test_export_detailed_results5_including_convex()

		if not os.path.isfile(self.detailed5b_export):
			self.test_export_detailed_results5b_including_convex()

		# Import dataframe for first worksheet
		df_lines = pd.read_excel(io=self.detailed5_export)  # type: pd.DataFrame
		df_cont_command = pd.read_excel(io=self.detailed5b_export)  # type: pd.DataFrame

		# Compare DataFrames to confirm they match
		# Compares the first column
		self.assertTrue(df_lines.iloc[:,0].equals(df_cont_command.iloc[:,0]))

		# Compares the second column where only expect small differences due to rounding errors
		col_num = 1
		df = pd.concat([df_lines.iloc[:,col_num], df_cont_command.iloc[:,col_num]]).drop_duplicates(keep=False)
		self.assertTrue(len(df) == 2)

		# Compares the third column where only expect small differences due to rounding errors
		col_num = 2
		df = pd.concat([df_lines.iloc[:,col_num], df_cont_command.iloc[:,col_num]]).drop_duplicates(keep=False)
		self.assertTrue(len(df) == 2)

		# Compares the third column where only expect small differences due to rounding errors
		col_num = 4
		df = pd.concat([df_lines.iloc[:,col_num], df_cont_command.iloc[:,col_num]]).drop_duplicates(keep=False)
		self.assertTrue(len(df) == 0)


class TestCombineMultiple(unittest.TestCase):
	"""
		Class to test that combining multiple runs works as expected
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

		self.cls_extract = MockExtractResults()

	def test_combine_single_results_set(self):
		""" Tests exporting of a single results set works """

		src_paths = (self.results1, )

		df, extract_vars = self.cls_extract.combine_multiple_runs(self=self.cls_extract, search_paths=src_paths)

		self.assertTrue(df.shape, (10,72))
		self.assertTrue(len(extract_vars), 6)
		self.assertTrue(pscharmonics.constants.PowerFactory.pf_x1 in extract_vars)

