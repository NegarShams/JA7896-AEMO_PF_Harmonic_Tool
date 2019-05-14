import unittest
import os
import pandas as pd
import sys

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
SEARCH_PTH = os.path.join(TESTS_DIR, 'results1')
HAST_INPUTS = os.path.join(SEARCH_PTH, 'HAST Inputs_test.xlsx')
HAST_Results1 = os.path.join(SEARCH_PTH, 'FS_SC1_Base_Case.csv')

RESULTS_EXTRACT_1 = os.path.join(TESTS_DIR, 'Processed Hast Results1.xlsx')
RESULTS_EXTRACT_2 = os.path.join(TESTS_DIR, 'Processed Hast Results2.xlsx')
RESULTS_EXTRACT_3 = os.path.join(TESTS_DIR, 'Processed Hast Results3.xlsx')

two_up = os.path.abspath(os.path.join(TESTS_DIR, '../..'))
sys.path.append(two_up)
import Process_HAST_extract as TestModule
import hast2_1.constants as constants

# If full test then will confirm that the importing of the variables from the hast file is correct but the
# testing for this is done elsewhere and this takes longer to run.  Setting to false skips the longer tests.
FULL_TEST = True

# ----- UNIT TESTS -----
class TestStandAloneFunctions(unittest.TestCase):
	"""
		Test class for all standalone functions
	"""
	@classmethod
	def setUpClass(cls):
		"""
			Setup the test class with parameters that are used in several different tests
		:return:
		"""
		# Dictionary of terminals used in test
		cls.test_dict_terminals = {('Bracetown.ElmSubstat', '220 kV A2.ElmTerm'): 'Bracetown 220 kV',
								   ('Clonee.ElmSubstat', '220 kV A1.ElmTerm'): 'Clonee 220 kV',
								   ('Clonee.ElmSubstat', '110 kV A1.ElmTerm'): 'Clonee 110 kV',
								   ('Kellis.ElmSubstat', '220 kV A1.ElmTerm'): 'Kellis 220 kV',
								   ('Maynooth.ElmSubstat', '220 kV A2.ElmTerm'): 'Maynooth 220 kV A',
								   ('Maynooth.ElmSubstat', '220 kV B1.ElmTerm'): 'Maynooth 220 kV B'
								   }
		# Variables as part of HAST_test_inputs for testing
		cls.test_vars = ['m:Z', 'c:Z_12']

		# Imports HAST inputs for further testing but full testing of import done elsewhere
		if FULL_TEST:
			cls.hast_file = TestModule.get_hast_values(search_pth=SEARCH_PTH)

	def test_extract_var_name(self):
		"""
			Tests the exctract var name function works correctly
		"""
		# Terminal name to test as input
		test_term_name = (r'\david.IntUser\AIM 2017-MODEL-07022019-TAP_TEST.IntPrj\Network Model.IntPrjfolder'
						r'\Network Data.IntPrjfolder\EirGrid.ElmNet\Bracetown.ElmSubstat\220 kV A2.ElmTerm')
		test_term_name2 = (r'\david.IntUser\AIM 2017-MODEL-07022019-TAP_TEST.IntPrj\Network Model.IntPrjfolder'
						  r'\Network Data.IntPrjfolder\EirGrid.ElmNet\Bracetown.ElmSubstat\110 kV A2.ElmTerm')

		# Mutual name to test as input
		test_mut_name = (r'\david.IntUser\AIM 2017-MODEL-07022019-TAP_TEST.IntPrj\Network Model.IntPrjfolder'
						 r'\Network Data.IntPrjfolder\EirGrid.ElmNet\ElmMut19_04_04_14_41_06'
						 r'\Bracetown 220 kV_Clonee 220 kV.ElmMut')
		test_mut_name2 = (r'\david.IntUser\AIM 2017-MODEL-07022019-TAP_TEST.IntPrj\Network Model.IntPrjfolder'
						 r'\Network Data.IntPrjfolder\EirGrid.ElmNet\ElmMut19_04_04_14_41_06'
						 r'\Clonee 220 kV_Bracetown 220 kV.ElmMut')
		test_mut_name3 = (r'\david.IntUser\AIM 2017-MODEL-07022019-TAP_TEST.IntPrj\Network Model.IntPrjfolder'
						  r'\Network Data.IntPrjfolder\EirGrid.ElmNet\ElmMut19_04_04_14_41_06'
						  r'\Clonee 220 kV_Clonee 110 kV.ElmMut')
		test_mut_name4 = (r'\user1.IntUser\AIM 2017-MODEL-07022019-TAP_CloneePh3.IntPrj\Network Model.IntPrjfolder'
						  r'\Network Data.IntPrjfolder\EirGrid.ElmNet\ElmMut19_04_15_01_54_25'
						  r'\Kellis 220 kV_Maynooth 220 kV A.ElmMut')
		test_mut_name5 = (r'\user1.IntUser\AIM 2017-MODEL-07022019-TAP_CloneePh3.IntPrj\Network Model.IntPrjfolder'
						  r'\Network Data.IntPrjfolder\EirGrid.ElmNet\ElmMut19_04_15_01_54_25'
						  r'\Maynooth 220 kV A_Maynooth 220 kV B.ElmMut')

		# Main function tests
		# Test that terminal extraction works correctly
		new_var, ref_term = TestModule.extract_var_name(var_name=test_term_name,
														dict_of_terms=self.test_dict_terminals)
		self.assertEqual(new_var, 'Bracetown 220 kV')

		# Test that mutual extaction works correctly
		new_var, ref_term = TestModule.extract_var_name(var_name=test_mut_name,
														dict_of_terms=self.test_dict_terminals)
		self.assertEqual(new_var, ('Bracetown 220 kV_Clonee 220 kV', 'Clonee 220 kV_Bracetown 220 kV'))
		self.assertEqual(ref_term, ('Bracetown 220 kV', 'Clonee 220 kV'))

		# Test that mutual extraction works correctly the other way around
		new_var, ref_term = TestModule.extract_var_name(var_name=test_mut_name2,
														dict_of_terms=self.test_dict_terminals)
		self.assertEqual(new_var, ('Clonee 220 kV_Bracetown 220 kV', 'Bracetown 220 kV_Clonee 220 kV'))
		self.assertEqual(ref_term, ('Clonee 220 kV', 'Bracetown 220 kV'))

		# Test that mutual extraction works correctly when testing nodes which are not defined
		new_var, ref_term = TestModule.extract_var_name(var_name=test_mut_name3,
														dict_of_terms=self.test_dict_terminals)
		self.assertEqual(new_var, ('Clonee 220 kV_Clonee 110 kV', 'Clonee 110 kV_Clonee 220 kV'))
		self.assertEqual(ref_term, ('Clonee 220 kV', 'Clonee 110 kV'))

		# Test raises an error if terminal has not been defined in dictionary
		self.assertRaises(KeyError,
						  TestModule.extract_var_name,
						  test_term_name2, self.test_dict_terminals)

		# Test works for different format of variable name (test_mut_name4)
		new_var, ref_term = TestModule.extract_var_name(var_name=test_mut_name4,
														dict_of_terms=self.test_dict_terminals)
		self.assertEqual(new_var, ('Kellis 220 kV_Maynooth 220 kV A', 'Maynooth 220 kV A_Kellis 220 kV'))
		self.assertEqual(ref_term, ('Kellis 220 kV', 'Maynooth 220 kV A'))

		# Test works for different format of variable name (test_mut_name5)
		new_var, ref_term = TestModule.extract_var_name(var_name=test_mut_name5,
														dict_of_terms=self.test_dict_terminals)
		self.assertEqual(new_var, ('Maynooth 220 kV A_Maynooth 220 kV B', 'Maynooth 220 kV B_Maynooth 220 kV A'))
		self.assertEqual(ref_term, ('Maynooth 220 kV A', 'Maynooth 220 kV B'))

	def test_extract_var_type(self):
		"""
			Tests the extract var type function works correctly
		"""
		# Variable types for testing
		test_var_type = 'c:R_12 in Ohms'
		test_var_type2 = 'c:R_12'
		test_var_type3 = 'b:fnow in Hz'
		test_var_type4 = 'An other sort of value'

		# Conduct tests of different input arrangements
		res_type = TestModule.extract_var_type(test_var_type)
		self.assertEqual(res_type,'c:R_12')
		res_type = TestModule.extract_var_type(test_var_type2)
		self.assertEqual(res_type, 'c:R_12')

		# Check suitably detects a frequency input
		res_type = TestModule.extract_var_type(test_var_type3)
		self.assertEqual(res_type, 'b:fnow')

		# Ensure fails if input is a value that is not a correct variable type
		self.assertRaises(IOError, TestModule.extract_var_type, test_var_type4)

	def test_process_single_file(self):
		"""
			Tests the imported HAST results file can be returned in a suitable dataframe
		"""
		# Import file to obtain dataframe
		df = TestModule.process_file(pth_file=HAST_Results1, hast_inputs=self.hast_file)
		# Confirm it is the correct dimensions, properties and values
		self.assertEqual(df.shape[0], 396)
		self.assertEqual(df.shape[1], 15)
		self.assertEqual(df.columns.levels[0][0], 'Bracetown 220 kV')
		# Confirm that mutual impedance data is being added in
		self.assertEqual(df.columns.levels[1][1], 'Bracetown 220 kV_Clonee 220 kV')
		self.assertEqual(df.columns.levels[1][3], 'Clonee 220 kV_Bracetown 220 kV')
		# Confirm naming is correct
		self.assertEqual(df.columns.names[0], 'Terminal')
		# Check a single value
		self.assertAlmostEqual(df.iloc[5,10], 78.484256, places=5)

	def test_process_r12_file(self):
		"""
			Tests that a HAST output containing R12 and X12 data can be handled correctly
		"""
		r12_test_file = os.path.join(TESTS_DIR, 'results4','FS_SC1_Base_case.csv')
		df = TestModule.process_file(pth_file=r12_test_file, hast_inputs=self.hast_file)
		self.assertEqual(df.shape,(196, 15))
		self.assertEqual(df.columns.levels[6][1], 'c:R_12')
		self.assertAlmostEqual(df.iloc[5,10], 78.484256, places=5)

	def test_combine_multiple_files(self):
		"""
			Test to import and combine multiple HAST results files and export them
		"""
		combined_df = TestModule.import_all_results(search_pth=SEARCH_PTH, hast_inputs=self.hast_file)
		# Confirm size correct
		self.assertEqual(combined_df.shape, (396,30))
		# Confirm columns correct
		self.assertEqual(combined_df.columns.levels[1][1],'Bracetown 220 kV_Clonee 220 kV')
		self.assertEqual(combined_df.columns.names[5],constants.ResultsExtract.lbl_FullName)
		# Check a random value
		self.assertAlmostEqual(combined_df.iloc[10,15],29.087961, places=4)

	def test_combine_r12_files(self):
		"""
			Tests that importing results with r12 and x12 data works correctly
		"""
		search_path = os.path.join(TESTS_DIR, 'results4')
		combined_df = TestModule.import_all_results(search_pth=search_path, hast_inputs=self.hast_file)
		# Confirm size correct
		self.assertEqual(combined_df.shape, (196,90))
		# Confirm r12 and x12 headers correct
		result_type_labels = combined_df.columns.levels[6]
		self.assertTrue('c:R_12' in result_type_labels)
		self.assertTrue('c:X_12' in result_type_labels)
		# Check a random value
		self.assertAlmostEqual(combined_df.iloc[10,15], 29.087961, places=4)

	@unittest.skipIf(not FULL_TEST, 'Tests that import HAST file skipped')
	def test_obtaining_vars_from_hast(self):
		"""
			Function tests importing the hast file to obtain the variables for export
			The full HAST analysis dict import is tested elsewhere
		"""
		if self.hast_file is None:
			self.hast_file = TestModule.get_hast_values(search_pth=SEARCH_PTH)

		vars_to_export = self.hast_file.vars_to_export()
		self.assertEqual(vars_to_export, self.test_vars)

	def test_hast_import_failure(self):
		"""
			Function checks that if no hast input files exist then will error
		"""
		test_search_pth = os.path.join(TESTS_DIR, 'Results_Fail_Test')
		self.assertRaises(IOError,
						  TestModule.get_hast_values,
						  test_search_pth)

	def test_export_multiple_files(self):
		"""
			Test to export the imported and combined results, just tests that the file exists rather than
			the contents of it.
		"""
		# File to export to
		target_file = RESULTS_EXTRACT_1
		# Check if file already exists and delete
		if os.path.isfile(target_file):
			os.remove(target_file)

		combined_df = TestModule.import_all_results(search_pth=SEARCH_PTH, hast_inputs=self.hast_file)
		TestModule.extract_results(pth_file=target_file, df=combined_df, vars_to_export=self.test_vars)
		# Confirm newly created file exists
		self.assertTrue(os.path.isfile(target_file))

	def test_export_multiple_r12_files(self):
		"""
			Tests that exporting r12 and x12 data works, just tests that the file exists rather than content of it
		"""
		# File to export to
		target_file = RESULTS_EXTRACT_1
		search_path = os.path.join(TESTS_DIR, 'results4')
		test_vars = ['m:R', 'm:X', 'm:Z', 'c:R_12', 'c:X_12', 'c:Z_12']
		# Check if file already exists and delete
		if os.path.isfile(target_file):
			os.remove(target_file)

		combined_df = TestModule.import_all_results(search_pth=search_path, hast_inputs=self.hast_file)
		TestModule.extract_results(pth_file=target_file, df=combined_df, vars_to_export=test_vars)
		# Confirm newly created file exists
		self.assertTrue(os.path.isfile(target_file))

	def test_export_multiple_files_without_graphs(self):
		"""
			Test to export the imported and combined results without graph plotting, just tests that the file exists rather than
			the contents of it.
		"""
		# File to export to
		target_file = RESULTS_EXTRACT_1
		# Check if file already exists and delete
		if os.path.isfile(target_file):
			os.remove(target_file)

		combined_df = TestModule.import_all_results(search_pth=SEARCH_PTH, hast_inputs=self.hast_file)
		TestModule.extract_results(pth_file=target_file, df=combined_df, vars_to_export=self.test_vars, plot_graphs=False)
		# Confirm newly created file exists
		self.assertTrue(os.path.isfile(target_file))

	def test_split_contingency_filter_values(self):
		"""
			Confirm the for a variety of filter names they are extracted correctly
		"""
		# Set to None to avoid test module error flagging
		test_cont_name = None
		file_splits = None

		name = dict()
		name['FS_SC1_CDU2-CRU2-ckt1_F1_500.0Hz_20.0MVAR'] = ('CDU2-CRU2-ckt1', 'F1_500.0Hz_20.0MVAR')
		name['FS_SC1_Base_Case'] = ('Base_Case','')
		name['FS_SC1_Base_Case_F1_550.0Hz_5.0MVAR'] = ('Base_Case', 'F1_550.0Hz_5.0MVAR')
		name['FS_SC1_WOO4_T4204_F1_500.0Hz_5.0MVAR'] = ('WOO4_T4204', 'F1_500.0Hz_5.0MVAR')
		name['FS_SC1_WOO4_T4204'] = ('WOO4_T4204','')

		for sc, (test_cont_name, test_filter_name) in name.items():
			file_splits = sc.split('_')
			cont_name, filter_name = TestModule.split_contingency_filter_values(list_of_terms=file_splits)
			self.assertEqual(cont_name, test_cont_name)
			self.assertEqual(filter_name, test_filter_name)

		cont_name, filter_name = TestModule.split_contingency_filter_values(list_of_terms=file_splits,
																			starting_point=3)
		self.assertNotEqual(cont_name, test_cont_name)

	def test_grouping(self):
		"""
			Test that if dataframe is passed it will return the correct grouping numbers
		"""
		# Create test dataframe
		data = [[1, 1, 1], [2, 2, 2], [3, 3, 3]]
		df = pd.DataFrame(data)

		cols = [['ab', 'ab', 'ab'], ['abc', 'bac', 'abc'], ['1', '2', '3']]
		index = pd.MultiIndex.from_arrays(cols, names=('A', 'B', 'C'))
		df.columns = index
		# Confirm if multiple results then are split by 2nd level
		df_grouping = TestModule.graph_grouping(df=df, group_by=('A','B'))
		self.assertEqual(df_grouping[0],('ab_abc', 2))
		self.assertEqual(df_grouping[1], ('ab_bac', 1))

		# Confirm if only 1 result then they are moved up a level
		df_grouping = TestModule.graph_grouping(df=df, group_by=('A', 'C'))
		self.assertEqual(df_grouping[0], ('ab', 3))

		# Confirm if only 1 result then they are moved up a level
		df_grouping = TestModule.graph_grouping(df=df, group_by=('A',))
		self.assertEqual(df_grouping[0], ('ab', 3))

	def test_split_plots(self):
		"""
			Test that it is suitable extracted
		"""
		max_plots = 5
		start_col=1
		graph_grouping = [('a',15),('b',1),('c',4)]
		graphs = TestModule.split_plots(max_plots=max_plots,
										start_col=start_col,
										graph_groups=graph_grouping)
		keys = list(graphs.keys())
		self.assertEqual(keys[0],'a(1)')
		self.assertEqual(keys[3],'b')
		self.assertEqual(graphs['c'][2],19)


class TestFunctionsWithoutHASTinputs(unittest.TestCase):
	"""
		Carries out tests which do not require the HAST inputs to be loaded
	"""
	def test_extract_var_name2(self):
		"""
			Tests the extract var name function works correctly
		"""
		test_dict_terminals = {('Carrigdangan.ElmSubstat', '110 kV A1.ElmTerm'): 'Carrigdangan_110 kV A1',
							   ('Dunmanway.ElmSubstat', '110 kV A1.ElmTerm'): 'Dunmanway_110 kV A1',
							   ('Boggeragh.ElmSubstat', '110 kV A1.ElmTerm'): 'Boggeragh_110 kV A1'
							   }
		# Mutual name to test as input
		test_mut_name = (r'\zuelli_r.IntUser\AIM 2019-CP0930_Carrigdangan.IntPrj\Network Model.IntPrjfolder'
						 r'\Network Data.IntPrjfolder\EirGrid.ElmNet\ElmMut19_05_13_10_18_27'
						 r'\Dunmanway_110 kV A1_Boggeragh_110 kV A1.ElmMut')

		# Main function tests
		# Test that terminal extraction works correctly
		new_var, ref_term = TestModule.extract_var_name(var_name=test_mut_name,
														dict_of_terms=test_dict_terminals)
		self.assertEqual(new_var, ('Dunmanway_110 kV A1_Boggeragh_110 kV A1',
								   'Boggeragh_110 kV A1_Dunmanway_110 kV A1'))

@unittest.skipIf(not FULL_TEST, 'Slower tests have been skipped')
class TestImportingMultipleResults(unittest.TestCase):
	"""
		Class to test importing and combining multiple results into a single export
	"""
	@classmethod
	def setUpClass(cls):
		"""
			Setup the test class with parameters that are used in several different tests
		:return:
		"""
		# Search paths that will be used for combining results
		cls.search_pth_1 = os.path.join(TESTS_DIR, 'results1')
		cls.search_pth_2 = os.path.join(TESTS_DIR, 'results2')
		cls.search_pth_3 = os.path.join(TESTS_DIR, 'results3')
		cls.search_pth_4 = os.path.join(TESTS_DIR, 'results4')

	def test_with_single_hast_file(self):
		"""
			Tests that both import methods produce the same results
		"""
		# DataFrames should be equal
		# Method 1 = Using a list of inputs
		df1, vars_to_export = TestModule.combine_multiple_hast_runs(
			search_pths=[self.search_pth_1]
		)

		# Method 2 = defining single folder and terminals for lookup
		hast_file = TestModule.get_hast_values(search_pth=self.search_pth_1)
		df2 = TestModule.import_all_results(search_pth=self.search_pth_1, hast_inputs=hast_file)

		# Sort dataframes to that they are in same order
		df1.sort_index(axis=1, level=0, inplace=True)
		df2.sort_index(axis=1, level=0, inplace=True)
		# Check if both imported dataframes are equal
		self.assertTrue(df1.equals(df2))
		self.assertEqual(vars_to_export, hast_file.vars_to_export())

	def test_multiple_hast_imports_duplicate_datasets(self):
		"""
			Tests the actual importing of multiple files but with same datasets
		"""
		df, _ = TestModule.combine_multiple_hast_runs(
			search_pths=[self.search_pth_1,
						 self.search_pth_2]
		)
		self.assertEqual(df.shape[1], 34)
		self.assertAlmostEqual(df.iloc[250, 20], -120.48696100000001, places=4)

	def test_multiple_hast_imports_different_datasets(self):
		"""
			Tests the actual importing of multiple files but with different datasets
		"""
		df, _ = TestModule.combine_multiple_hast_runs(
			search_pths=[self.search_pth_1,
						 self.search_pth_3]
		)
		self.assertEqual(df.shape[1], 58)
		self.assertAlmostEqual(df.iloc[250,25], -123.121829, places=4)

	def test_multiple_hast_imports_do_not_drop_duplicates(self):
		"""
			Tests the actual importing of multiple files but with dropping of duplicates
		"""
		df, _ = TestModule.combine_multiple_hast_runs(
			search_pths=[self.search_pth_1,
						 self.search_pth_2],
			drop_duplicates=False
		)
		self.assertEqual(df.shape[1], 60)
		self.assertAlmostEqual(df.iloc[250,25], -123.121829, places=4)

	def test_multiple_hast_imports_different_length_datasets(self):
		"""
			Tests the actual importing of multiple files but with different length datasets
		"""
		df, _ = TestModule.combine_multiple_hast_runs(
			search_pths=[self.search_pth_1,
						 self.search_pth_4]
		)
		self.assertEqual(df.shape[1], 120)
		self.assertAlmostEqual(df.iloc[120,25], 58.063766, places=4)

	def test_extract_multiple_files(self):
		"""
			Function imports multiple results and extracts them to a single results file
			NB:  Only tests that the file is created and no the contents of it
			Does not test that the import is correct since that is tested elsewhere
		"""
		# File to export to
		target_file = RESULTS_EXTRACT_2
		# Check if file already exists and delete
		if os.path.isfile(target_file):
			os.remove(target_file)

		df, vars_to_export = TestModule.combine_multiple_hast_runs(
			search_pths=[self.search_pth_1,
						 self.search_pth_3]
		)

		TestModule.extract_results(pth_file=target_file, df=df, vars_to_export=vars_to_export)
		# Confirm newly created file exists
		self.assertTrue(os.path.isfile(target_file))

	def test_extract_multiple_files_different_data_lengths(self):
		"""
			Function imports multiple results and extracts them to a single results file
			NB:  Only tests that the file is created and no the contents of it
			Does not test that the import is correct since that is tested elsewhere
		"""
		# File to export to
		target_file = RESULTS_EXTRACT_3
		# Check if file already exists and delete
		if os.path.isfile(target_file):
			os.remove(target_file)

		df, vars_to_export = TestModule.combine_multiple_hast_runs(
			search_pths=[self.search_pth_1,
						 self.search_pth_4]
		)

		TestModule.extract_results(pth_file=target_file, df=df, vars_to_export=vars_to_export)
		# Confirm newly created file exists
		self.assertTrue(os.path.isfile(target_file))



