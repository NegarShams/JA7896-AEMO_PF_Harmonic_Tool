import unittest
import os
import shutil

from .context import pscharmonics
import pscharmonics.pf as TestModule


# If full test then will confirm that the importing of the variables from the hast file is correct but the
# testing for this is done elsewhere and this takes longer to run.  Setting to false skips the longer tests.
FULL_TEST = True
TESTS_DIR = os.path.dirname(os.path.abspath(__file__))

TestModule.DEBUG_MODE=True


class TestPFInitialisation(unittest.TestCase):
	""" Tests that the correct python version can be found """


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
		results_file = TestModule.main(import_workbook=hast_inputs_file,
									   results_export_folder=self.results_export_folder,
									   uid='stage0',
									   include_nom_voltage=False)
		self.assertTrue(os.path.isfile(results_file))

	def test_stage0_processing_v220(self):
		"""
			Produces stage 0 results that are needed for the harmonic limits calculation
			v220 - This now takes into consideration the nominal voltage of the node being investigated
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage0.xlsx')
		results_file = TestModule.main(import_workbook=hast_inputs_file,
									   results_export_folder=self.results_export_folder_v220,
									   uid='stage0_v220')
		self.assertTrue(os.path.isfile(results_file))

	def test_stage1_processing(self):
		"""
			Produces stage 1 results that are needed for the harmonic limits calculation
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage1.xlsx')
		results_file = TestModule.main(import_workbook=hast_inputs_file,
									   results_export_folder=self.results_export_folder,
									   uid='stage1',
									   include_nom_voltage=False)
		self.assertTrue(os.path.isfile(results_file))

	def test_stage1_processing_v220(self):
		"""
			Produces stage 1 results that are needed for the harmonic limits calculation
			v220 - This now takes into consideration the nominal voltage of the node being investigated
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage1.xlsx')
		results_file = TestModule.main(import_workbook=hast_inputs_file,
									   results_export_folder=self.results_export_folder_v220,
									   uid='stage1_v220')
		self.assertTrue(os.path.isfile(results_file))

	def test_stage2_processing(self):
		"""
			Produces stage 2 results that are needed for the harmonic limits calculation
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage2.xlsx')
		results_file = TestModule.main(import_workbook=hast_inputs_file,
									   results_export_folder=self.results_export_folder,
									   uid='stage2',
									   include_nom_voltage=False)
		self.assertTrue(os.path.isfile(results_file))

	def test_stage2_processing_v220(self):
		"""
			Produces stage 2 results that are needed for the harmonic limits calculation
			v220 - This now takes into consideration the nominal voltage of the node being investigated
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage2.xlsx')
		results_file = TestModule.main(import_workbook=hast_inputs_file,
									   results_export_folder=self.results_export_folder_v220,
									   uid='stage2_v220')
		self.assertTrue(os.path.isfile(results_file))

	def test_stage_all_processing_v220(self):
		"""
			Produces set of results for all stages in single HAST run
			This is a combination of stage0, stage1 and stage2 HAST
			input files
			v220 - This now takes into consideration the nominal voltage of the node being investigated
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage_all.xlsx')
		results_file = TestModule.main(import_workbook=hast_inputs_file,
									   results_export_folder=self.results_export_folder_v220,
									   uid='stage_all_v220')
		self.assertTrue(os.path.isfile(results_file))

	def test_results1_v2_0_processing(self):
		"""
			Produces a new set of HAST results based on update to include nominal voltage in v2.2.0
		:return:
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'results1', 'HAST Inputs_test.xlsx')
		results_file = TestModule.main(import_workbook=hast_inputs_file,
									   results_export_folder=TESTS_DIR,
									   uid='results1_v220')
		self.assertTrue(os.path.isfile(results_file))

	def test_results5_production(self):
		"""
			Produces a new set of HAST results using a HAST inputs file which contains different
			nominal voltages
		:return:
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST Inputs_test_different_voltages.xlsx')
		results_file = TestModule.main(import_workbook=hast_inputs_file,
									   results_export_folder=TESTS_DIR,
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
		results_file = TestModule.main(import_workbook=hast_inputs_file,
									   results_export_folder=r'D:\Test Fail\Newly created\not possible',
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

	def test_non_convergent_base_case(self):
		"""
			Checks that if there is a non-convergent base case for all studies it fails
		"""
		# Nominal voltage not included
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_non_convergent.xlsx')
		self.assertRaises(RuntimeError, TestModule.main,
						  import_workbook=hast_inputs_file,
						  results_export_folder=self.results_export_folder,
						  uid='non_convergent')

	def test_one_convergent_base_case(self):
		"""
			Checks that if there is a single non-convergent base case then it continues, returning a single error
		"""
		# Nominal voltage not included
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_partially_convergent.xlsx')
		results_file = TestModule.main(import_workbook=hast_inputs_file,
									   results_export_folder=TESTS_DIR,
									   uid='partially_convergent')
		self.assertTrue(os.path.isfile(results_file))

class TestHASTInputsProcessing(unittest.TestCase):
	"""
		Class of tests to confirm the input processing is working correctly
	"""
	def test_long_terminal_names(self):
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_long_term_names.xlsx')
		with TestModule.hast2.file_io.Excel(print_info=print, print_error=print) as excel_cls:
			analysis_dict = excel_cls.import_excel_harmonic_inputs(
				pth_workbook=hast_inputs_file)
		with self.assertRaises(ValueError):
			TestModule.hast2.file_io.HASTInputs(hast_inputs=analysis_dict)

	def test_duplicated_study_case_names(self):
		"""
			Test confirms that if a HAST inputs file is used which contains duplicated study case names
			then a critical error will be raised
		:return:
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_duplicated_study_cases.xlsx')
		with TestModule.hast2.file_io.Excel(print_info=print, print_error=print) as excel_cls:
			analysis_dict = excel_cls.import_excel_harmonic_inputs(
				pth_workbook=hast_inputs_file)
		with self.assertRaises(ValueError):
			TestModule.hast2.file_io.HASTInputs(hast_inputs=analysis_dict,
													  filename=hast_inputs_file)

	def test_duplicated_contingency_names(self):
		"""
			Test confirms that if a HAST inputs file is used which contains duplicated contingency case names
			then a critical error will be raised
		:return:
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_duplicated_contingencies.xlsx')
		with TestModule.hast2.file_io.Excel(print_info=print, print_error=print) as excel_cls:
			analysis_dict = excel_cls.import_excel_harmonic_inputs(
				pth_workbook=hast_inputs_file)
		with self.assertRaises(ValueError):
			TestModule.hast2.file_io.HASTInputs(hast_inputs=analysis_dict,
													  filename=hast_inputs_file)

	def test_duplicated_terminals(self):
		"""
			Test confirms that if a HAST inputs file is used which contains duplicated terminals names
			then a critical error will be raised
		:return:
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_duplicated_terminals.xlsx')
		with TestModule.hast2.file_io.Excel(print_info=print, print_error=print) as excel_cls:
			analysis_dict = excel_cls.import_excel_harmonic_inputs(
				pth_workbook=hast_inputs_file)
		with self.assertRaises(ValueError):
			TestModule.hast2.file_io.HASTInputs(hast_inputs=analysis_dict,
													  filename=hast_inputs_file)

	def test_duplicated_filters(self):
		"""
			Test confirms that if a HAST inputs file is used which contains duplicated filter names
			then a critical error will be raised
		:return:
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_duplicated_filters.xlsx')
		with TestModule.hast2.file_io.Excel(print_info=print, print_error=print) as excel_cls:
			analysis_dict = excel_cls.import_excel_harmonic_inputs(
				pth_workbook=hast_inputs_file)
		with self.assertRaises(ValueError):
			TestModule.hast2.file_io.HASTInputs(hast_inputs=analysis_dict,
													  filename=hast_inputs_file)

	def test_missing_study_case_names(self):
		"""
			Confirm that if a study case name is missing it is skipped but
			suitable warning is raised.
			# TODO: Detect warning created
		:return:
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_missing_study_case_names.xlsx')
		with TestModule.hast2.file_io.Excel(print_info=print, print_error=print) as excel_cls:
			analysis_dict = excel_cls.import_excel_harmonic_inputs(
				pth_workbook=hast_inputs_file)
		# Get handle to hast
		cls_hast = TestModule.hast2.file_io.HASTInputs(hast_inputs=analysis_dict)
		# Confirm that length is only 1
		self.assertTrue(len(cls_hast.sc_names) == 1)

	# TODO: Test not yet implemented
	@unittest.skip('Not Yet Implemented')
	def test_new_hast_import(self):
		"""
			Test confirms that if a HAST inputs file is used which contains duplicated terminals names
			then a critical error will be raised
		:return:
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage0.xlsx')
		TestModule.hast2.file_io.HASTInputs(filename=hast_inputs_file)
