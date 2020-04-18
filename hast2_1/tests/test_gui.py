import unittest
import os
import hast2_1.gui as test_module

FULL_TEST = True

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
SEARCH_PTH = os.path.join(TESTS_DIR, 'results1')

def mock_load_settings_file(self):
	"""
		Mock function that is run when the study settings file is selected
	:return:
	"""
	# TODO: May want to include something here to determine further functionality testing
	# If importing workbook was a success then change the state of the other buttons and coninue
	return None


class TestGui(unittest.TestCase):
	"""
		UnitTest package to confirm that GUI can be produced correctly
	"""
	def test_file_selection(self):
		"""Tests for the file selection function"""
		self.assertRaises(SyntaxError, test_module.file_selector)
		self.assertRaises(SyntaxError, test_module.file_selector, open_file=True, save_dir=True)
		# Following tests require the GUI to load and so are only run occasionally
		if FULL_TEST:
			print('\n ## Select cancel for UnitTest of file dialog## \n')
			target_file = test_module.file_selector(initial_pth=TESTS_DIR,
													open_file=True)
			self.assertEqual(target_file, '')

	@unittest.skipIf(not FULL_TEST, 'Skipping GUI creation for testing')
	def test_main_gui(self):
		"""Tests for the main GUI"""
		print('** Click cancel **')
		gui = test_module.MainGUI(start_directory=TESTS_DIR)
		self.assertEqual(gui.results_files_list, [])

	@unittest.skipIf(not FULL_TEST, 'Skipping GUI creation for testing')
	def test_main_gui2(self):
		"""Tests for the main GUI"""
		# Mock normal function for testing
		# TODO: Better way to test this functionality
		original_function = test_module.MainGui.load_settings_file
		test_module.MainGui.load_settings_file = mock_load_settings_file
		gui = test_module.MainGui(start_directory=TESTS_DIR)
		test_module.MainGui.load_settings_file = original_function
		self.assertEqual(gui.abort, True)
