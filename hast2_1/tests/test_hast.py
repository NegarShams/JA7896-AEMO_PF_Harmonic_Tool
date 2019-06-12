import unittest
import HAST_V2_1 as TestModule
import os
import shutil

# If full test then will confirm that the importing of the variables from the hast file is correct but the
# testing for this is done elsewhere and this takes longer to run.  Setting to false skips the longer tests.
FULL_TEST = True
TESTS_DIR = os.path.dirname(os.path.abspath(__file__))

TestModule.DEBUG_MODE=True


# ----- UNIT TESTS -----
class TestHast(unittest.TestCase):
	@classmethod
	def setUp(cls):
		cls.results_export_folder = os.path.join(TESTS_DIR, 'STAGES')
		cls.results_export_folder_v220 = os.path.join(TESTS_DIR, 'STAGES_v220')
		if not os.path.isdir(cls.results_export_folder):
			os.mkdir(cls.results_export_folder)

	def test_stage0_processing(self):
		"""
			Produces stage 0 results that are needed for the harmonic limits calculation
		"""
		# Nominal voltage not included
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage0.xlsx')
		results_file = TestModule.main(Import_Workbook=hast_inputs_file,
									   Results_Export_Folder=self.results_export_folder,
									   uid='stage0',
									   include_nom_voltage=False)
		self.assertTrue(os.path.isfile(results_file))

	def test_stage0_processing_v220(self):
		"""
			Produces stage 0 results that are needed for the harmonic limits calculation
			v220 - This now takes into consideration the nominal voltage of the node being investigated
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage0.xlsx')
		results_file = TestModule.main(Import_Workbook=hast_inputs_file,
									   Results_Export_Folder=self.results_export_folder_v220,
									   uid='stage0_v220')
		self.assertTrue(os.path.isfile(results_file))

	def test_stage1_processing(self):
		"""
			Produces stage 1 results that are needed for the harmonic limits calculation
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage1.xlsx')
		results_file = TestModule.main(Import_Workbook=hast_inputs_file,
									   Results_Export_Folder=self.results_export_folder,
									   uid='stage1',
									   include_nom_voltage=False)
		self.assertTrue(os.path.isfile(results_file))

	def test_stage1_processing_v220(self):
		"""
			Produces stage 1 results that are needed for the harmonic limits calculation
			v220 - This now takes into consideration the nominal voltage of the node being investigated
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage1.xlsx')
		results_file = TestModule.main(Import_Workbook=hast_inputs_file,
									   Results_Export_Folder=self.results_export_folder_v220,
									   uid='stage1_v220')
		self.assertTrue(os.path.isfile(results_file))

	def test_stage2_processing(self):
		"""
			Produces stage 2 results that are needed for the harmonic limits calculation
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage2.xlsx')
		results_file = TestModule.main(Import_Workbook=hast_inputs_file,
									   Results_Export_Folder=self.results_export_folder,
									   uid='stage2',
									   include_nom_voltage=False)
		self.assertTrue(os.path.isfile(results_file))

	def test_stage2_processing_v220(self):
		"""
			Produces stage 2 results that are needed for the harmonic limits calculation
			v220 - This now takes into consideration the nominal voltage of the node being investigated
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage2.xlsx')
		results_file = TestModule.main(Import_Workbook=hast_inputs_file,
									   Results_Export_Folder=self.results_export_folder_v220,
									   uid='stage2_v220')
		self.assertTrue(os.path.isfile(results_file))

	def test_results1_v2_0_processing(self):
		"""
			Produces a new set of HAST results based on update to include nominal voltage in v2.2.0
		:return:
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'results1', 'HAST Inputs_test.xlsx')
		results_file = TestModule.main(Import_Workbook=hast_inputs_file,
									   Results_Export_Folder=TESTS_DIR,
									   uid='results1_v220')
		self.assertTrue(os.path.isfile(results_file))

	def test_results5_production(self):
		"""
			Produces a new set of HAST results using a HAST inputs file which contains different
			nominal voltages
		:return:
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST Inputs_test_different_voltages.xlsx')
		results_file = TestModule.main(Import_Workbook=hast_inputs_file,
									   Results_Export_Folder=TESTS_DIR,
									   uid='results5(diff_voltages)')
		# results file should = False since Export to Excel = False in inputs spreadsheet
		self.assertFalse(os.path.isfile(results_file))

	def test_path_not_found(self):
		"""
			Tests that if a path is provided that cannot be created the script still runs and a folder is created
			elsewhere
		:return:
		"""
		test_uid = 'results(path_not_found)'
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage0.xlsx')
		results_file = TestModule.main(Import_Workbook=hast_inputs_file,
									   Results_Export_Folder=r'D:\Test Fail\Newly created\not possible',
									   uid=test_uid)
		# results file should = False since Export to Excel = False in inputs spreadsheet
		self.assertTrue(os.path.isfile(results_file))
		path_to_raw_results = os.path.join(os.path.dirname(results_file), test_uid)
		self.assertTrue(os.path.isdir(path_to_raw_results))

		# Tidy up by deleting path and results file (shutil deletes a non-empty tree)
		shutil.rmtree(path_to_raw_results)
		os.remove(results_file)
		# Confirm deleted successfully
		self.assertFalse(os.path.isfile(results_file))
		self.assertFalse(os.path.isdir(path_to_raw_results))


class TestHASTInputsProcessing(unittest.TestCase):
	"""
		Class of tests to confirm the input processing is working correctly
	"""
	def test_long_terminal_names(self):
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_long_term_names.xlsx')
		with TestModule.hast2.excel_writing.Excel(print_info=print, print_error=print) as excel_cls:
			analysis_dict = excel_cls.import_excel_harmonic_inputs(
				workbookname=hast_inputs_file)
		with self.assertRaises(ValueError):
			TestModule.hast2.excel_writing.HASTInputs(hast_inputs=analysis_dict)

