"""
	Script will deal with contain all the power factory functions used for HAST to carry out and produce results

	UNIT TESTING
		This script contains unit tests and these should be run whenever a change is made to confirm everything is
		still functioning as expected.
		Some unittests will require a working instance of PowerFactory on the test machine
"""

import unittest

# TODO: Need to conisder if this is the best apporach to importing powerfactory or if should use from parent
import powerfactory

class PowerFactory:
	""" Class stores all the handles to the power factory functions so that they can be defined in one location
		only
	"""
	def __init__(self, app):
		"""
			Initialise
		:param powerfactory.GetApplication app: Handle to powerfactory app
		"""
		# These may not actually be required yet
		self.ldf = app.GetFromStudyCase("ComLdf")  # Get load flow command
		self.hldf = app.GetFromStudyCase("ComHldf")  # Get Harmonic load flow
		self.frq = app.GetFromStudyCase("ComFsweep")  # Get Frequency Sweep Command
		self.ini = app.GetFromStudyCase("ComInc")  # Get Dynamic Initialisation
		self.sim = app.GetFromStudyCase("ComSim")  # Get Dynamic Simulation
		self.shc = app.GetFromStudyCase("ComShc")  # Get short circuit command
		self.res = app.GetFromStudyCase("ComRes")  # Get Result Export Command
		self.wr = app.GetFromStudyCase("ComWr")  # Get Write command for wmf and bmp files

	def check_parallel(self):
		""" Function confirms whether PowerFactory has a Parallel Computing Manager setup such that it can be used to
			run the calculations in parallel.  If it does not then this will need to be setup and an appropriate
			message will be sent to the user
		"""
		

# ----------------------- UNIT TESTS  ------------------------------------------