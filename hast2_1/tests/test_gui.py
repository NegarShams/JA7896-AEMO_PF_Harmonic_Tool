import unittest
import os
import hast2_1.gui as TestModule

FULL_TEST = True

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
SEARCH_PTH = os.path.join(TESTS_DIR, 'results1')


class TestGui(unittest.TestCase):
	"""
		UnitTest package to confirm that GUI can be produced correctly
	"""
	def test_file_selection(self):
		"""Tests for the file selection function"""
		self.assertRaises(SyntaxError, TestModule.file_selector)
		self.assertRaises(SyntaxError, TestModule.file_selector, open_file=True, save_dir=True)
		# Following tests require the GUI to load and so are only run occasionally
		if FULL_TEST:
			print('\n ## Select cancel for UnitTest of file dialog## \n')
			target_file = TestModule.file_selector(initial_pth=TESTS_DIR,
												   open_file=True)
			self.assertEqual(target_file, '')

	@unittest.skipIf(not FULL_TEST, 'Skipping GUI creation for testing')
	def test_main_gui(self):
		"""Tests for the main GUI"""
		print('** Click cancel **')
		gui = TestModule.MainGUI(start_directory=TESTS_DIR)
		self.assertEqual(gui.results_files_list, [])