import unittest
import os
import sys
import shutil

from .context import pscharmonics
import pscharmonics.pf as TestModule


# If full test then will confirm that the importing of the variables from the hast file is correct but the
# testing for this is done elsewhere and this takes longer to run.  Setting to false skips the longer tests.
FULL_TEST = True
TESTS_DIR = os.path.join(os.path.dirname(__file__), 'test_files')

pf_test_model = 'pscharmonics_test_model'

# When this is set to True tests which require initialising PowerFactory are skipped
skip_slow_tests = True

TestModule.DEBUG_MODE=True


class TestPFInitialisation(unittest.TestCase):
	""" Tests that the correct python version can be found and then PowerFactory can be initialised """
	@classmethod
	def setUp(cls):
		""" Gets initial references that are need for PowerFactory class initialisation """
		cls.pf = pscharmonics.pf.PowerFactory()

	def test_adding_python_paths_works(self):
		"""
			This test only confirms that PowerFactory can be imported, there is no actual test to be
			performed
		:return:
		"""
		self.pf.add_python_paths()

		# Confirm that python is now in the path
		self.assertTrue(any(['PowerFactory' in x for x in sys.path]))

	@unittest.skipUnless(skip_slow_tests, 'Tests that require initialising PowerFactory have been skipped')
	def test_pf_initialisation(self):
		""" Function tests that powerfactory can be initialised """
		self.pf.initialise_power_factory()

		# Test confirms that app has been initialised and that the installation directory matches the system path
		# that has been provided
		pf_directory = pscharmonics.pf.app.GetInstallationDirectory()

		self.assertEqual(os.path.abspath(pf_directory), os.path.abspath(self.pf.c.dig_path))

	@unittest.skipUnless(skip_slow_tests, 'Tests that require initialising PowerFactory have been skipped')
	def test_pf_processors(self):
		""" Function tests checking of powerfactory parallel processors """
		self.pf.initialise_power_factory()
		self.pf.check_parallel_processing()


@unittest.skipUnless(skip_slow_tests, 'Tests that require initialising PowerFactory have been skipped')
class TestsOnPFCase(unittest.TestCase):
	"""
		This class carries out tests on a specific PowerFactory case, if it doesn't exist in the PowerFactory model
		already then it is initially imported and then removed at the end
	"""
	pf = None
	pf_test_project = None

	@classmethod
	def setUpClass(cls):
		""" Initialise PowerFactory and then check if model already exists """
		cls.pf = pscharmonics.pf.PowerFactory()
		cls.pf.initialise_power_factory()

		# Try to activate the test project and if it doesn't work then load power factory case in
		if not cls.pf.active_project(project_name=pf_test_model):
			pf_test_file = os.path.join(TESTS_DIR, '{}.pfd'.format(pf_test_model))

			# Import the project
			cls.pf.import_project(project_pth=pf_test_file)
			cls.pf.active_project(project_name=pf_test_model)

		cls.pf_test_project = cls.pf.get_active_project()

	def test_deactivate_project(self):
		""" Function tests that activating and deactivating a project works as expected """
		# Confirm project already active
		pf_prj = self.pf.active_project(project_name=pf_test_model)

		# Confirm that if project is activated
		self.assertEqual(pf_prj, self.pf.get_active_project())

		# Deactivate project
		self.pf.deactivate_project()

		# Confirm no longer equal and that project can be deactivated
		self.assertNotEqual(pf_prj, self.pf.get_active_project())
		self.assertTrue(self.pf.get_active_project() is None)

	@classmethod
	def tearDownClass(cls):
		""" Function ensures the deletion of the PowerFactory project """
		# Deactivate and then delete the project
		cls.pf.deactivate_project()
		cls.pf.delete_object(pf_obj=cls.pf_test_project)

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
