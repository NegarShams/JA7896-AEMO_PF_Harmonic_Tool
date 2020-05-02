"""
	Context file to ensure that irrelevant of the installation method the test and package files can
	be found
"""

import os
import sys

# Add package path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Import main module
import pscharmonics