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
import pscharmonics.batch_mode as batch_mode

# Reload all modules so that if run from PowerFactory doesn't need to be closed and reopened during debugging
logger = importlib.reload(logger)
file_io = importlib.reload(file_io)
pf = importlib.reload(pf)
constants = importlib.reload(constants)
gui = importlib.reload(gui)
batch_mode = importlib.reload(batch_mode)

if constants.logger is None:
	constants.logger = logger.Logger()
	# Redirect exceptions to be capture by logger
	sys.excepthook = constants.logger.exception_handler