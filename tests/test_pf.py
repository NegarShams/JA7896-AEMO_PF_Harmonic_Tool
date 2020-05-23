import unittest
import os
import sys
import shutil
import pandas as pd
import math
import random
import string

from tests.context import pscharmonics
import pscharmonics.pf as TestModule


# If full test then will confirm that the importing of the variables from the hast file is correct but the
# testing for this is done elsewhere and this takes longer to run.  Setting to false skips the longer tests.
FULL_TEST = True
TESTS_DIR = os.path.join(os.path.dirname(__file__), 'test_files')

pf_test_project = 'pscharmonics_test_model'
pf_test_sc = 'High Load Case'
pf_test_os = 'HighLoadTap_testing'
pf_test_inputs = 'Inputs.xlsx'

# When this is set to True tests which require initialising PowerFactory are skipped
include_slow_tests = True

TestModule.DEBUG_MODE=True

# OLD tests are being skipped
include_old_tests = False

# Set to True and created excel outputs will be deleted during test run
test_delete_excel_outputs = False

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

	@unittest.skipUnless(include_slow_tests, 'Tests that require initialising PowerFactory have been skipped')
	def test_pf_initialisation(self):
		""" Function tests that powerfactory can be initialised """
		self.pf.initialise_power_factory()

		# Test confirms that app has been initialised and that the installation directory matches the system path
		# that has been provided
		pf_directory = pscharmonics.pf.app.GetInstallationDirectory()

		self.assertEqual(os.path.abspath(pf_directory), os.path.abspath(self.pf.c.dig_path))

@unittest.skipUnless(include_slow_tests, 'Tests that require initialising PowerFactory have been skipped')
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
		if not cls.pf.activate_project(project_name=pf_test_project):
			pf_test_file = os.path.join(TESTS_DIR, '{}.pfd'.format(pf_test_project))

			# Import the project
			cls.pf.import_project(project_pth=pf_test_file)
			cls.pf.activate_project(project_name=pf_test_project)

		cls.pf_test_project = cls.pf.get_active_project()

	def test_deactivate_project(self):
		""" Function tests that activating and deactivating a project works as expected """
		# Confirm project already active
		pf_prj = self.pf.activate_project(project_name=pf_test_project)

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


@unittest.skipUnless(include_slow_tests, 'Tests that require initialising PowerFactory have been skipped')
class TestPFProject(unittest.TestCase):
	"""
		Tests PF Project functions
	"""
	@classmethod
	def setUpClass(cls):
		""" Initialise PowerFactory and then check if model already exists """
		cls.pf = pscharmonics.pf.PowerFactory()
		cls.pf.initialise_power_factory()

		# Try to activate the test project and if it doesn't work then load power factory case in
		if not cls.pf.activate_project(project_name=pf_test_project):
			pf_test_file = os.path.join(TESTS_DIR, '{}.pfd'.format(pf_test_project))

			# Import the project
			cls.pf.import_project(project_pth=pf_test_file)
			cls.pf.activate_project(project_name=pf_test_project)

		cls.pf_test_project = cls.pf.get_active_project()

		# Create DataFrame for project tests
		cls.test_name = 'TEST'
		data = [cls.test_name, pf_test_project, pf_test_sc, pf_test_os]
		columns = pscharmonics.constants.StudySettings.studycase_columns
		cls.df = pd.DataFrame(data=data).transpose()
		cls.df.columns = columns
		# Set the index to be based on the unique name
		cls.df.set_index(pscharmonics.constants.StudySettings.name, inplace=True, drop=False)

	def test_create_project_temporary_folders(self):
		"""
			Confirm that new project instances can be created
		:return None:
		"""
		# Create new project instances
		uid = 'TEST_CASE'
		pf_projects = pscharmonics.pf.create_pf_project_instances(df_study_cases=self.df, uid=uid)

		# Check defined as expected
		pf_project = pf_projects[pf_test_project]
		self.assertEqual(pf_test_project, pf_project.name)
		self.assertTrue(pf_project.prj_active)

		# Confirm temporary folders exist in the temporary project folder location
		temp_folder_location = pscharmonics.pf.app.GetProjectFolder('study')
		folder = temp_folder_location.GetContents(
			'{}_{}.{}'.format(
				pscharmonics.constants.PowerFactory.temp_sc_folder,
				uid,
				pscharmonics.constants.PowerFactory.pf_folder_type
			)
		)
		# Confirm folder exists
		self.assertTrue(len(folder) > 0)
		# Confirm objects match with first element
		self.assertEqual(folder[0], pf_project.sc_folder)

		# Delete folders and then confirm folder no longer exists
		pf_project.delete_temp_folders()
		folder = temp_folder_location.GetContents(
			'{}_{}.{}'.format(
				pscharmonics.constants.PowerFactory.temp_sc_folder,
				uid,
				pscharmonics.constants.PowerFactory.pf_folder_type
			)
		)
		self.assertTrue(len(folder) == 0)

		# Deactivate project
		pf_project.project_state(deactivate=True)
		# Confirm deactivated
		active_project = self.pf.get_active_project()
		self.assertTrue(active_project is None)

	def test_create_task_auto(self):
		"""
			Confirm that new project instances can be created and that it contains the auto_executable task
		:return None:
		"""
		# Create new project instances
		uid = 'TEST_CASE'
		pf_projects = pscharmonics.pf.create_pf_project_instances(df_study_cases=self.df, uid=uid)

		# Check defined as expected
		pf_project = pf_projects[pf_test_project]
		self.assertEqual(pf_test_project, pf_project.name)
		self.assertTrue(pf_project.prj_active)

		# Confirm temporary folders exist in the temporary project folder location
		task_auto = pf_project.sc_folder.GetContents(
			'{}_{}.{}'.format(
				pscharmonics.constants.General.cmd_autotasks_leader,
				uid,
				pscharmonics.constants.PowerFactory.autotasks_command
			)
		)
		# Confirm folder exists
		self.assertTrue(len(task_auto) > 0)
		# Confirm objects match with first element
		self.assertEqual(task_auto[0], pf_project.task_auto)

		# Change uid and create new task_auto
		uid='TEST_TASK_AUTO'
		pf_project.uid = uid
		task_auto = pf_project.create_task_auto()

		# Confirm that task_auto command now contains new uid and doesn't match previous task_auto
		self.assertTrue(uid in str(task_auto))
		self.assertNotEqual(task_auto, pf_project.task_auto)


		# Delete temporary folders
		pf_project.delete_temp_folders()

	def test_copy_studycases(self):
		""" tests that new study cases can be created """
		# Create new project instances
		uid = 'TEST_CASE'
		df_test_project = self.df[self.df[pscharmonics.constants.StudySettings.name]==self.test_name]
		pf_project = pscharmonics.pf.PFProject(name=pf_test_project, df_studycases=df_test_project, uid=uid)

		# Confirm study cases created
		self.assertTrue(self.test_name in pf_project.base_sc.keys())

		# Confirm can activate study case
		sc = pf_project.base_sc[self.test_name]
		# Confirm initially deactivated, change state and then confirm active
		self.assertFalse(sc.active)
		sc.toggle_state()
		self.assertTrue(sc.active)

		# Confirm can deactivate
		sc.toggle_state(deactivate=True)
		self.assertFalse(sc.active)

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()

	def test_studycase_lf_assignment(self):
		""" Tests that new study cases can be created with approapriate load flow settings """
		# Load flow settings
		def_inputs_file = os.path.join(TESTS_DIR, 'Inputs.xlsx')
		with pd.ExcelFile(def_inputs_file) as wkbk:
			# Import here should match pscconsulting.file_io.StudyInputsDev().process_lf_settings
			df = pd.read_excel(
				wkbk,
				sheet_name=pscharmonics.constants.HASTInputs.lf_settings,
				usecols=(3,), skiprows=3, header=None, squeeze=True
			)

		# Create instance with complete set of settings
		lf_settings = pscharmonics.file_io.LFSettings(
			existing_command=df.iloc[0], detailed_settings=df.iloc[1:])

		# Create new project instances
		uid = 'TEST_CASE'
		df_test_project = self.df[self.df[pscharmonics.constants.StudySettings.name]==self.test_name]
		pf_project = pscharmonics.pf.PFProject(
			name=pf_test_project, df_studycases=df_test_project, uid=uid,
			lf_settings=lf_settings
		)

		# Confirm can activate study case
		sc = pf_project.base_sc[self.test_name]

		# Activate study case
		sc.toggle_state()
		self.assertTrue(sc.active)

		# Run load flow and should get error code 0 returned
		self.assertTrue(pscharmonics.constants.General.cmd_leader in str(sc.ldf))
		self.assertEqual(sc.ldf.Execute(), 0)

		# Run load flow using in built class command
		self.assertTrue(sc.run_load_flow())
		self.assertTrue(sc.ldf_convergent)

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()

	def test_studycase_fs_assignment(self):
		""" Tests that new study cases can be created with approapriate load flow settings """
		# Load flow settings
		def_inputs_file = os.path.join(TESTS_DIR, 'Inputs.xlsx')
		with pd.ExcelFile(def_inputs_file) as wkbk:
			# Import here should match pscconsulting.file_io.StudyInputsDev().process_lf_settings
			df = pd.read_excel(
				wkbk,
				sheet_name=pscharmonics.constants.HASTInputs.fs_settings,
				usecols=(3,), skiprows=3, header=None, squeeze=True
			)

		# Create instance with complete set of settings
		fs_settings = pscharmonics.file_io.FSSettings(
			existing_command=df.iloc[0], detailed_settings=df.iloc[1:])

		# Create new project instances
		uid = 'TEST_CASE'
		df_test_project = self.df[self.df[pscharmonics.constants.StudySettings.name]==self.test_name]
		pf_project = pscharmonics.pf.PFProject(
			name=pf_test_project, df_studycases=df_test_project, uid=uid,
			fs_settings=fs_settings
		)

		# Confirm can activate study case
		sc = pf_project.base_sc[self.test_name]

		# Activate study case
		sc.toggle_state()
		self.assertTrue(sc.active)

		# Although no load flow created can run using default, and then confirm correctly executed
		self.assertFalse(sc.fs is None)
		self.assertEqual(sc.fs.Execute(), 0)
		# TODO: Confirm settings match inputs

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()

	def test_studycase_fs_result_export(self):
		""" Tests that new study cases can be created with a frequency scan and then export the results
			to a suitable file
		"""
		test_export_pth = os.path.join(TESTS_DIR)

		# Load flow settings
		def_inputs_file = os.path.join(TESTS_DIR, 'Inputs.xlsx')
		with pd.ExcelFile(def_inputs_file) as wkbk:
			# Import here should match pscconsulting.file_io.StudyInputsDev().process_lf_settings
			df = pd.read_excel(
				wkbk,
				sheet_name=pscharmonics.constants.HASTInputs.fs_settings,
				usecols=(3,), skiprows=3, header=None, squeeze=True
			)

		# Create instance with complete set of settings
		fs_settings = pscharmonics.file_io.FSSettings(
			existing_command=df.iloc[0], detailed_settings=df.iloc[1:])

		# Create new project instances
		uid = 'TEST_CASE'
		df_test_project = self.df[self.df[pscharmonics.constants.StudySettings.name]==self.test_name]
		pf_project = pscharmonics.pf.PFProject(
			name=pf_test_project, df_studycases=df_test_project, uid=uid,
			fs_settings=fs_settings
		)

		# Get handle for study case
		sc = pf_project.base_sc[self.test_name]
		# Activate study case
		sc.toggle_state()
		self.assertTrue(sc.active)

		# Set results path for associated study case and check file doesn't already exist
		sc.res_pth = test_export_pth
		test_export_file = os.path.join(test_export_pth, '{}.csv'.format(sc.name))
		if os.path.isfile(test_export_file):
			os.remove(test_export_file)
		# Create results export command
		export_cmd, res_pth = sc.set_results_export(result=sc.results)

		# Confirm returned path matches expected value
		self.assertEqual(test_export_file, res_pth)

		# Run frequency scan and then export results
		# Although no load flow created can run using default, and then confirm correctly executed
		sc.fs.Execute()
		# Confirm returns 0 for successful execute
		self.assertEqual(export_cmd.Execute(), 0)
		# Confirm file exists and then delete
		self.assertTrue(os.path.isfile(test_export_file))
		os.remove(test_export_file)

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()

	def test_studycase_complete_studies_creation(self):
		""" Tests that new study cases can be created with relevant studies created in a single
			command
		"""
		test_export_pth = os.path.join(TESTS_DIR)

		# Load flow settings
		def_inputs_file = os.path.join(TESTS_DIR, 'Inputs.xlsx')
		with pd.ExcelFile(def_inputs_file) as wkbk:
			# Import here should match pscconsulting.file_io.StudyInputsDev().process_lf_settings
			df = pd.read_excel(
				wkbk,
				sheet_name=pscharmonics.constants.HASTInputs.fs_settings,
				usecols=(3,), skiprows=3, header=None, squeeze=True
			)

			# Create instance with complete set of settings
			fs_settings = pscharmonics.file_io.FSSettings(
				existing_command=df.iloc[0], detailed_settings=df.iloc[1:])

			# Import here should match pscconsulting.file_io.StudyInputsDev().process_lf_settings
			df = pd.read_excel(
				wkbk,
				sheet_name=pscharmonics.constants.HASTInputs.lf_settings,
				usecols=(3,), skiprows=3, header=None, squeeze=True
			)

			# Create instance with complete set of settings
			lf_settings = pscharmonics.file_io.LFSettings(
				existing_command=df.iloc[0], detailed_settings=df.iloc[1:])

		# Create new project instances
		uid = 'TEST_CASE'
		df_test_project = self.df[self.df[pscharmonics.constants.StudySettings.name]==self.test_name]
		pf_project = pscharmonics.pf.PFProject(
			name=pf_test_project, df_studycases=df_test_project, uid=uid
		)

		# Get handle for study case
		sc = pf_project.base_sc[self.test_name]
		# Activate study case
		sc.toggle_state()
		self.assertTrue(sc.active)

		# Set results path for associated study case and check file doesn't already exist
		sc.res_pth = test_export_pth
		test_export_file = os.path.join(test_export_pth, '{}.csv'.format(sc.name))
		if os.path.isfile(test_export_file):
			os.remove(test_export_file)

		# Create studies
		sc.create_studies(lf_settings=lf_settings, fs_settings=fs_settings)

		# Confirm returned path matches expected value
		self.assertEqual(test_export_file, sc.fs_result_exports[0])

		# Run load flow, frequency scan and export
		self.assertEqual(sc.ldf.Execute(), 0)
		self.assertEqual(sc.fs.Execute(), 0, msg='Check PQ license exists since frequency scan failed')
		self.assertEqual(sc.fs_export_cmd.Execute(), 0)

		self.assertTrue(os.path.isfile(test_export_file))
		os.remove(test_export_file)

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()

	@classmethod
	def tearDownClass(cls):
		""" Function ensures the deletion of the PowerFactory project """
		# Deactivate and then delete the project
		cls.pf.deactivate_project()
		cls.pf.delete_object(pf_obj=cls.pf_test_project)


@unittest.skipUnless(include_slow_tests, 'Tests that require initialising PowerFactory have been skipped')
class TestPFProjectContingencyCases(unittest.TestCase):
	"""
		Tests creation of contingency cases and fault events within the PFProject class
	"""
	# Names used for testing contingencies (convergent and non-convergent)
	test_cont = 'TEST Cont'
	test_cont_nc = 'TEST Cont NC'

	@classmethod
	def setUpClass(cls):
		""" Initialise PowerFactory and then check if model already exists """
		cls.pf = pscharmonics.pf.PowerFactory()
		cls.pf.initialise_power_factory()

		# Try to activate the test project and if it doesn't work then load power factory case in
		if not cls.pf.activate_project(project_name=pf_test_project):
			pf_test_file = os.path.join(TESTS_DIR, '{}.pfd'.format(pf_test_project))

			# Import the project
			cls.pf.import_project(project_pth=pf_test_file)
			cls.pf.activate_project(project_name=pf_test_project)

		cls.pf_test_project = cls.pf.get_active_project()

		# Create DataFrame for project tests
		cls.test_name = 'TEST'
		data = [cls.test_name, pf_test_project, pf_test_sc, pf_test_os]
		columns = pscharmonics.constants.StudySettings.studycase_columns
		cls.df = pd.DataFrame(data=data).transpose()
		cls.df.columns = columns
		# Set the index to be based on the unique name
		cls.df.set_index(pscharmonics.constants.StudySettings.name, inplace=True, drop=False)

		# Import all settings
		def_inputs_file = os.path.join(TESTS_DIR, 'Inputs.xlsx')
		cls.settings = pscharmonics.file_io.StudyInputsDev(pth_file=def_inputs_file)

	def test_create_fault_cases(self):
		"""
			Tests that fault cases can be created for a contingency command
		"""
		# Create new project instances
		uid = 'TEST_CASE'
		fc_name = 'TEST Cont'
		df_test_project = self.df[self.df[pscharmonics.constants.StudySettings.name]==self.test_name]
		pf_project = pscharmonics.pf.PFProject(
			name=pf_test_project, df_studycases=df_test_project, uid=uid
		)

		# Create fault events
		fault_cases = pf_project.create_fault_cases(contingencies=self.settings.contingencies)
		fc = fault_cases[fc_name]

		switch_name = 'Switch_213211'
		self.assertEqual(fc.loc_name, fc_name)
		event = fc.GetContents('{}.{}'.format(switch_name, pscharmonics.constants.PowerFactory.pf_switch_event))[0]
		self.assertTrue(switch_name in str(event.p_target))
		self.assertTrue(event.i_switch==0)

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()

	def test_create_results_files_works_correctly(self):
		"""
			Tests that appropriate results files are created
		"""
		# Create new project instances
		uid = 'TEST_CASE'
		df_test_project = self.df[self.df[pscharmonics.constants.StudySettings.name]==self.test_name]
		pf_project = pscharmonics.pf.PFProject(
			name=pf_test_project, df_studycases=df_test_project, uid=uid
		)

		# Get handle to test study case
		sc = pf_project.base_sc[self.test_name]

		# Confirm no results files have been defined initially
		self.assertTrue(sc.fs_results is None)
		self.assertTrue(sc.cont_results is None)

		# Create the results files and confirm they now reference a DataObject
		sc.create_results_files()

		search_string = pscharmonics.constants.PowerFactory.pf_results
		self.assertTrue(search_string in str(sc.fs_results))
		self.assertTrue(search_string in str(sc.cont_results))

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()

	def test_create_cont_analysis_5_fault_cases(self):
		"""
			Tests that fault cases can be created and then a suitable contingency command
			also created.
		:return:
		"""
		# Create new project instances
		uid = 'Test_fault_analysis'
		df_test_project = self.df[self.df[pscharmonics.constants.StudySettings.name]==self.test_name]
		pf_project = pscharmonics.pf.PFProject(
			name=pf_test_project, df_studycases=df_test_project, uid=uid
		)
		# Get handle to test study case
		sc = pf_project.base_sc[self.test_name]


		# Create fault events
		fault_cases = pf_project.create_fault_cases(contingencies=self.settings.contingencies)

		# Create results files for contingency analysis
		sc.create_results_files()
		# Create load flow command
		sc.create_load_flow(lf_settings=None)
		# Create contingency based on fault cases
		sc.create_cont_analysis(fault_cases=fault_cases)

		# Carry out validation by firstly confirming the the cont_analysis function has been created and is the
		# correct type
		self.assertTrue(pscharmonics.constants.PowerFactory.pf_cont_analysis in str(sc.cont_analysis))


		# Then confirm that the sc.df index contains the default contingencies
		self.assertTrue(self.test_cont in sc.df.index)
		self.assertTrue(self.test_cont_nc in sc.df.index)

		# Ensure study case is active
		sc.toggle_state()
		self.assertEqual(sc.cont_analysis.Execute(), 0)

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()

	def test_create_cont_analysis_non_existant_cmd(self):
		"""
			Tests that fault cases can be created and then a suitable contingency command
			also created.
		:return:
		"""
		# Create new project instances
		uid = 'Test_fault_analysis'
		df_test_project = self.df[self.df[pscharmonics.constants.StudySettings.name]==self.test_name]
		pf_project = pscharmonics.pf.PFProject(
			name=pf_test_project, df_studycases=df_test_project, uid=uid
		)
		# Get handle to test study case
		sc = pf_project.base_sc[self.test_name]

		# Create fault events
		fault_cases = pf_project.create_fault_cases(contingencies=self.settings.contingencies)

		# Confirm that if runs with a non-existent command still creates the relevant contingency analysis cases

		# Create contingency based on fault cases
		sc.create_load_flow(lf_settings=None)
		# Create results files for contingency analysis
		sc.create_results_files()
		# Create contingency based on fault cases
		sc.create_cont_analysis(fault_cases=fault_cases, cmd='A None Existent Command')

		# Carry out validation by firstly confirming the the cont_analysis function has been created and is the
		# correct type
		self.assertTrue(pscharmonics.constants.PowerFactory.pf_cont_analysis in str(sc.cont_analysis))

		# Then confirm that the sc.df index contains the default contingencies
		self.assertTrue('TEST Cont' in sc.df.index)
		self.assertTrue('TEST Cont NC' in sc.df.index)

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()

	def test_empty_cont_analysis_if_no_loadflow(self):
		"""
			Tests that fault cases fails if provided with no inputs
		:return:
		"""
		# Create new project instances
		uid = 'Test_fault_analysis'
		df_test_project = self.df[self.df[pscharmonics.constants.StudySettings.name]==self.test_name]
		pf_project = pscharmonics.pf.PFProject(
			name=pf_test_project, df_studycases=df_test_project, uid=uid
		)
		# Get handle to test study case
		sc = pf_project.base_sc[self.test_name]

		self.assertTrue(sc.cont_analysis is None)
		# Create contingency based on no inputs but shouldn't raise an error since load flow command
		# has not been defined yet
		sc.create_cont_analysis()
		# Should still be none
		self.assertTrue(sc.cont_analysis is None)

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()

	def test_create_cont_analysis_fails_with_no_input(self):
		"""
			Tests that fault cases fails if provided with no inputs
		:return:
		"""
		# Create new project instances
		uid = 'Test_fault_analysis'
		df_test_project = self.df[self.df[pscharmonics.constants.StudySettings.name]==self.test_name]
		pf_project = pscharmonics.pf.PFProject(
			name=pf_test_project, df_studycases=df_test_project, uid=uid
		)
		# Get handle to test study case
		sc = pf_project.base_sc[self.test_name]

		# Create contingency based on fault cases
		sc.create_load_flow(lf_settings=None)
		# Create load flow command first
		self.assertRaises(SyntaxError, sc.create_cont_analysis)

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()

	def test_create_cont_analysis_using_command(self):
		"""
			Tests that if contingency analysis command is provided then that is used for the analysis and the
			study case dataframe is updated appropriately
		:return:
		"""
		# Create new project instances
		uid = 'Test_fault_analysis'
		df_test_project = self.df[self.df[pscharmonics.constants.StudySettings.name]==self.test_name]
		pf_project = pscharmonics.pf.PFProject(
			name=pf_test_project, df_studycases=df_test_project, uid=uid
		)
		# Get handle to test study case
		sc = pf_project.base_sc[self.test_name]

		# Create results files for contingency analysis
		sc.create_results_files()
		# Create load flow command
		sc.create_load_flow(lf_settings=None)
		# Create contingency based on fault cases
		sc.create_cont_analysis(cmd='Contingency Analysis')

		# Confirm cases created correctly
		self.assertTrue(len(sc.df.index) == 1)

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()

	def test_process_cont_results(self):
		"""
			TODO: Being developed
		:return:
		"""
		# Create new project instances
		uid = 'Test_fault_analysis'
		df_test_project = self.df[self.df[pscharmonics.constants.StudySettings.name]==self.test_name]
		pf_project = pscharmonics.pf.PFProject(
			name=pf_test_project, df_studycases=df_test_project, uid=uid
		)
		# Get handle to test study case
		sc = pf_project.base_sc[self.test_name]


		# Create fault events
		fault_cases = pf_project.create_fault_cases(contingencies=self.settings.contingencies)

		# Create results files for contingency analysis
		sc.create_results_files()
		# Create load flow command
		sc.create_load_flow(lf_settings=None)
		# Create contingency based on fault cases
		sc.create_cont_analysis(fault_cases=fault_cases)

		# Active study case and run cont analysis
		sc.toggle_state()
		sc.cont_analysis.Execute()

		# Confirm that initial status is nan
		c = pscharmonics.constants.Contingencies
		self.assertTrue(math.isnan(sc.df.loc[self.test_cont, c.status]))
		self.assertTrue(math.isnan(sc.df.loc[self.test_cont_nc, c.status]))

		# Test results
		sc.process_cont_results()

		# Confirm that the status for one of the DataFrames is now non-convergent
		self.assertTrue(sc.df.loc[self.test_cont, c.status])
		self.assertFalse(sc.df.loc[self.test_cont_nc, c.status])

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()

	def test_pre_case_check_single_study_case(self):
		"""
			Test that when the pre-case check is run for the one study cases
			the dataframe that is returned matches what is expected.  This is
			only testing the pre case check, the creation of the study cases
			happens elsewhere.

			# Test is based on inputs
		:return:
		"""

		# Create new project instances
		uid = 'Test_fault_analysis'
		df_test_project = self.df[self.df[pscharmonics.constants.StudySettings.name]==self.test_name]
		pf_project = pscharmonics.pf.PFProject(
			name=pf_test_project, df_studycases=df_test_project, uid=uid
		)

		# Carry out pre-case check using fault cases
		df = pf_project.pre_case_check(contingencies=self.settings.contingencies)

		# Confirm that dataframe takes the correct form with the relevant headers for the columns
		names = df.index.names
		self.assertTrue(pscharmonics.constants.Contingencies.sc in names)

		# Confirm the contents of a single value matches expectations
		c = pscharmonics.constants.Contingencies
		sc_result = df.loc[(self.test_name, self.test_cont), :]
		self.assertEqual(sc_result[c.cont], self.test_cont)
		self.assertEqual(sc_result[c.status], True)

		# Confirm the non-convergent case is as expected
		c = pscharmonics.constants.Contingencies
		sc_result = df.loc[(self.test_name, self.test_cont_nc), :]
		self.assertEqual(sc_result[c.cont], self.test_cont_nc)
		self.assertEqual(sc_result[c.status], False)

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()

	def test_pre_case_check_multiple_study_cases(self):
		"""
			Test that when the pre-case check is run for the two study cases
			the dataframe that is returned matches what is expected.

		:return:
		"""
		# Test Names (detailed in Inputs spreadsheet)
		test_case1 = 'BASE'
		test_case2 = 'BASE(1)'


		# Create projects
		uid = 'TEST_CASE'
		pf_projects = pscharmonics.pf.create_pf_project_instances(
			df_study_cases=self.settings.cases,
			uid=uid
		)

		# Get single project
		pf_project = pf_projects[pf_test_project]

		# Carry out pre-case check using fault cases on what should be two different study_cases
		df = pf_project.pre_case_check(contingencies=self.settings.contingencies)

		# Confirm that dataframe takes the correct form with the relevant headers for the columns
		names = df.index.names
		self.assertTrue(pscharmonics.constants.Contingencies.sc in names)

		# Since project is duplicated expect results to be the same in both cases
		c = pscharmonics.constants.Contingencies
		sc_result = df.loc[(test_case1, self.test_cont), :]
		self.assertEqual(sc_result[c.cont], self.test_cont)
		self.assertEqual(sc_result[c.status], True)

		# Confirm the non-convergent case is as expected
		c = pscharmonics.constants.Contingencies
		sc_result = df.loc[(test_case2, self.test_cont), :]
		self.assertEqual(sc_result[c.cont], self.test_cont)
		self.assertEqual(sc_result[c.status], True)

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()

	def test_case_creation(self):
		"""
			Test that when the pre-case check is run for the two study cases
			the dataframe that is returned matches what is expected.

		:return:
		"""
		# Create random path for temporary files
		target_pth = os.path.join(
			TESTS_DIR, ''.join(random.choice(string.ascii_uppercase + string.digits) for _ in range(6))
		)
		if os.path.isdir(target_pth):
			raise SyntaxError('Random path {} already exists'.format(target_pth))
		else:
			os.mkdir(target_pth)


		# Create projects
		uid = 'TEST_CASE'
		pf_projects = pscharmonics.pf.create_pf_project_instances(
			df_study_cases=self.settings.cases,
			uid=uid
		)

		# Get single project
		pf_project = pf_projects[pf_test_project]

		# Carry out pre-case check using fault cases on what should be two different study_cases
		df = pf_project.pre_case_check(contingencies=self.settings.contingencies)

		num_convergent_cases = len(df[df[pscharmonics.constants.Contingencies.status]==True].index)

		# Create cases
		pf_project.create_cases(export_pth=target_pth, contingencies=self.settings.contingencies)

		self.assertEqual(len(pf_project.cases_to_run), num_convergent_cases)

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()


	def test_case_auto_exec_creation(self):
		"""
			Test that the task_auto that is created can be populated with the relevant commands
		:return:
		"""
		# Create random path for temporary files
		target_pth = os.path.join(
			TESTS_DIR, ''.join(random.choice(string.ascii_uppercase + string.digits) for _ in range(6))
		)
		if os.path.isdir(target_pth):
			raise SyntaxError('Random path {} already exists'.format(target_pth))
		else:
			os.mkdir(target_pth)

		# Create projects
		uid = 'TEST_CASE'
		pf_projects = pscharmonics.pf.create_pf_project_instances(
			df_study_cases=self.settings.cases,
			uid=uid
		)

		# Get single project
		pf_project = pf_projects[pf_test_project]

		# Carry out pre-case check using fault cases on what should be two different study_cases
		df = pf_project.pre_case_check(contingencies=self.settings.contingencies)

		# Get number of convergent cases which will then be used to determine the number of files that should exist
		# in the folder once study completed
		num_convergent_cases = len(df[df[pscharmonics.constants.Contingencies.status]==True].index)

		# Create cases
		pf_project.create_cases(export_pth=target_pth, contingencies=self.settings.contingencies)
		# Update the auto_exec command to contain details of all of these cases
		pf_project.update_auto_exec()

		# Confirm that the contents of the auto_exec command matches the number of study cases expected
		# and the correct number of commands
		# TODO: Check number of study cases and commands matches expectations for auto command

		# Execute command and confirm that the number of entries in the results folder is correct
		pf_project.project_state()
		pf_project.task_auto.Execute()

		# Find out how many results have been created
		num_results = len([name for name in os.listdir(target_pth) if os.path.isfile(os.path.join(target_pth, name))])

		# Confirm number of results matches 1 for each convergent cases
		self.assertEqual(num_convergent_cases, num_results)

		# Tidy up by deleting temporary project folders
		pf_project.delete_temp_folders()


	@classmethod
	def tearDownClass(cls):
		""" Function ensures the deletion of the PowerFactory project """
		# Deactivate and then delete the project
		cls.pf.deactivate_project()
		cls.pf.delete_object(pf_obj=cls.pf_test_project)


class TestPFSingleProjectUsingInputs(unittest.TestCase):
	"""
		This class contains tests that are carried out using the complete set of inputs as part of integration
		testing
	"""
	@classmethod
	def setUpClass(cls):
		""" Initialise PowerFactory and then check if model already exists """
		cls.pf = pscharmonics.pf.PowerFactory()
		cls.pf.initialise_power_factory()

		# Try to activate the test project and if it doesn't work then load power factory case in
		if not cls.pf.activate_project(project_name=pf_test_project):
			pf_test_file = os.path.join(TESTS_DIR, '{}.pfd'.format(pf_test_project))

			# Import the project
			cls.pf.import_project(project_pth=pf_test_file)
			cls.pf.activate_project(project_name=pf_test_project)

		cls.pf_test_project = cls.pf.get_active_project()

		# Import all settings
		def_inputs_file = os.path.join(TESTS_DIR, 'Inputs.xlsx')
		cls.settings = pscharmonics.file_io.StudyInputsDev(pth_file=def_inputs_file)

	def test_pre_case_check_for_all_projects(self):
		"""
			Function tests that running pre_case check for all projects correctly reports non-convergent cases
			and produces a suitable excel export
		:return:
		"""
		c = pscharmonics.constants.Contingencies

		# Target excel path
		excel_pth = os.path.join(TESTS_DIR, 'pre_case_check.xlsx')
		# Check doesn't exist initially and if it does then delete
		if os.path.isfile(excel_pth):
			os.remove(excel_pth)
		self.assertFalse(os.path.exists(excel_pth))

		# Create projects
		uid = 'TEST_CASE'
		pf_projects = pscharmonics.pf.create_pf_project_instances(
			df_study_cases=self.settings.cases,
			uid=uid
		)

		# Get DataFrame from pre_case check and also ask to create file
		df_summary = pscharmonics.pf.run_pre_case_checks(
			pf_projects=pf_projects,
			export_pth=excel_pth,
			contingencies=self.settings.contingencies
		)

		# Confirm DataFrame is expected length and number of non_convergent cases is as expected
		self.assertEqual(len(df_summary.index), 4)
		self.assertEqual(len(df_summary[df_summary[c.status]==False].index), 2)

		# Confirm that excel spreadsheet has been created (and delete if necessary)
		self.assertTrue(os.path.isfile(excel_pth))
		if test_delete_excel_outputs:
			os.remove(excel_pth)

	@classmethod
	def tearDownClass(cls):
		""" Function ensures the deletion of the PowerFactory project """
		# Deactivate and then delete the project
		cls.pf.deactivate_project()
		cls.pf.delete_object(pf_obj=cls.pf_test_project)

# TODO: Create pre-case check for multiple projects

# ----- UNIT TESTS OLD -----
@unittest.skipUnless(include_old_tests, 'Old tests skipped')
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

@unittest.skipUnless(include_old_tests, 'Old tests skipped')
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

	@unittest.skip('Not Yet Implemented')
	def test_new_hast_import(self):
		"""
			Test confirms that if a HAST inputs file is used which contains duplicated terminals names
			then a critical error will be raised
		:return:
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage0.xlsx')
		TestModule.hast2.file_io.HASTInputs(filename=hast_inputs_file)
