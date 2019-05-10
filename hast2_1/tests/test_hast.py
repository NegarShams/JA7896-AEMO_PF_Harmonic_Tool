import unittest
import HAST_V2_1 as TestModule
import os

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
		if not os.path.isdir(cls.results_export_folder):
			os.mkdir(cls.results_export_folder)

	def test_stage0_processing(self):
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage0.xlsx')
		results_file = TestModule.main(Import_Workbook=hast_inputs_file,
									   Results_Export_Folder=self.results_export_folder)
		self.assertTrue(os.path.isfile(results_file))

	def test_stage1_processing(self):
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage1.xlsx')
		results_file = TestModule.main(Import_Workbook=hast_inputs_file,
									   Results_Export_Folder=self.results_export_folder)
		self.assertTrue(os.path.isfile(results_file))

	def test_stage2_processing(self):
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_stage2.xlsx')
		results_file = TestModule.main(Import_Workbook=hast_inputs_file,
									   Results_Export_Folder=self.results_export_folder)
		self.assertTrue(os.path.isfile(results_file))

