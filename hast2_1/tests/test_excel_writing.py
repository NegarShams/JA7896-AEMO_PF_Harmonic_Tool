import unittest
import hast2_1.excel_writing as TestModule
import os

# If full test then will confirm that the importing of the variables from the hast file is correct but the
# testing for this is done elsewhere and this takes longer to run.  Setting to false skips the longer tests.
FULL_TEST = True
TESTS_DIR = os.path.dirname(os.path.abspath(__file__))

# ----- UNIT TESTS -----
class TestHASTInputs(unittest.TestCase):
	def test_hast_inputs_v1_2(self):
		"""
			unittest to check that still runs if an old HAST Inputs format is used
		:return:
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_old_format(v1_2).xlsx')
		with TestModule.Excel() as excel_cls:
			analysis_dict = excel_cls.import_excel_harmonic_inputs(workbookname=hast_inputs_file)  # Reads in the Settings from the spreadsheet

		cls_hast_inputs = TestModule.HASTInputs(hast_inputs=analysis_dict,
												uid_time='HAST Inputs v1_2')
		print(cls_hast_inputs.list_of_terms)
		self.assertTrue(cls_hast_inputs.list_of_terms[0][3])
	def test_hast_inputs_v2_1_3(self):
		"""
			unittest to check that still runs if an old HAST Inputs format is used
		:return:
		"""
		hast_inputs_file = os.path.join(TESTS_DIR, 'HAST_Inputs_old_format(v2_1_3).xlsx')
		with TestModule.Excel() as excel_cls:
			analysis_dict = excel_cls.import_excel_harmonic_inputs(workbookname=hast_inputs_file)  # Reads in the Settings from the spreadsheet

		cls_hast_inputs = TestModule.HASTInputs(hast_inputs=analysis_dict,
												uid_time='HAST Inputs v2_1_3')
		print(cls_hast_inputs.list_of_terms)
		self.assertFalse(cls_hast_inputs.list_of_terms[0][3])