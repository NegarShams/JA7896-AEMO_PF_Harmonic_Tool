"""
#######################################################################################################################
###													Tests associated with process.py								###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###																													###
#######################################################################################################################
"""

import unittest
import os
import sys

from tests.context import pscharmonics

# If full test then will confirm that the importing of the variables from the inputs file is correct but the
# testing for this is done elsewhere and this takes longer to run.  Setting to false skips the longer tests.
FULL_TEST = True
TESTS_DIR = os.path.join(os.path.dirname(__file__), 'test_files')


class TestCreateConvex(unittest.TestCase):
	""" Tests that passing R/X data will return ConvexHull around data points """

	def setUp(self):
		""" Creates a random data set """

	def test_convex(self):
		""" Tests can be created """
