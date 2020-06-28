"""
#######################################################################################################################
###											Initialisation															###
###		Script deals with initialising the pscharmonics scripts														###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###																													###
#######################################################################################################################
"""
import importlib
import sys

import pscharmonics.logger as logger
import pscharmonics.file_io as file_io
import pscharmonics.pf as pf
import pscharmonics.constants as constants
import pscharmonics.gui as gui

# Reload all modules
logger = importlib.reload(logger)
file_io = importlib.reload(file_io)
pf = importlib.reload(pf)
constants = importlib.reload(constants)
gui = importlib.reload(gui)

# import pscharmonics.processing as processing

if constants.logger is None:
	constants.logger = logger.Logger()
	# Redirect exceptions to be capture by logger
	sys.excepthook = constants.logger.exception_handler