"""
#######################################################################################################################
###													test_batch_mode.py												###
###		Test code for running unittests when operating in batch mode												###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###																													###
#######################################################################################################################
"""

import unittest
import os
import glob
import shutil

from tests.context import pscharmonics

# If full test then will confirm that the importing of the variables from the inputs file is correct but the
# testing for this is done elsewhere and this takes longer to run.  Setting to false skips the longer tests.
FULL_TEST = True
TESTS_DIR = os.path.join(os.path.dirname(__file__), 'test_files')

pf_test_project = 'pscharmonics_test_model'

# When this is set to True tests which require initialising PowerFactory are skipped
include_slow_tests = True

# Set to True and created excel outputs will be deleted during test run
test_delete_excel_outputs = True


class TestPFBatchRun(unittest.TestCase):
	"""
		Function runs the PSC Harmonics script exactly as it is run in Batch Mode to confirm operation
	"""
	pf = None  # type: pscharmonics.pf.PowerFactory
	pf_test_project = str()

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


	def test_batch_mode(self):
		"""
			Function runs batch mode test
		:return:
		"""
		# Set UID value
		pscharmonics.constants.uid = 'Test_Batch'

		# Create path for results from detailed tests
		target_export_pth = os.path.join(TESTS_DIR, 'Batch_Test')
		if os.path.isdir(target_export_pth):
			print('Existing contents in path: {} will be deleted'.format(target_export_pth))
			shutil.rmtree(target_export_pth)
		else:
			os.mkdir(target_export_pth)

		target_results_filename = 'Results_{}.xlsx'.format(pscharmonics.constants.uid)

		# Import settings for Detailed Study
		settings_file = os.path.join(TESTS_DIR, 'Inputs_Batch.xlsx')
		inputs = pscharmonics.file_io.StudyInputs(pth_file=settings_file)

		inputs.settings.export_folder = target_export_pth
		inputs.settings.results_name = target_results_filename


		# Run batch study and confirm a success
		success = pscharmonics.batch_mode.run(test_settings=inputs)

		self.assertTrue(success)

		# Confirm results files exists
		self.assertTrue(os.path.isdir(target_export_pth))

		# Expected results file
		expected_results_file = os.path.join(target_export_pth, target_results_filename)
		self.assertTrue(os.path.isfile(expected_results_file))


	@classmethod
	def tearDownClass(cls):
		""" Function ensures the deletion of the PowerFactory project """
		# Deactivate and then delete the project
		# cls.pf.deactivate_project()
		# cls.pf.delete_object(pf_obj=cls.pf_test_project)
