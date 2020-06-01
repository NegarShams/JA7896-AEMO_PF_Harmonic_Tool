"""
	Context file to ensure that irrelevant of the installation method the test and package files can
	be found.

	Other useful guide:  https://docs.python-guide.org/writing/structure/
"""

import os
import sys

# Add package path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))


# Import main module
# noinspection PyUnusedLocal
import pscharmonics

# Directory where all test files will be located
test_files_dir = os.path.join(os.path.dirname(__file__), 'test_files')