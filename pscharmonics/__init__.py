"""
#######################################################################################################################
###											Initialisation															###
###		Script deals with initialising the pscharmonics scripts														###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###																													###
#######################################################################################################################
"""
import pscharmonics.logger as logger
import pscharmonics.file_io as file_io
import pscharmonics.pf as pf
import pscharmonics.constants as constants
import pscharmonics.gui as gui
import pscharmonics.processing as processing
import sys

if constants.logger is None:
	constants.logger = logger.Logger()
	# Redirect exceptions to be capture by logger
	sys.excepthook = constants.logger.exception_handler