"""
#######################################################################################################################
###											PSC Harmonics															###
###		Script produced as part of JA7896 project for Automated Running of Frequency Scans in PowerFactory 			###
###																									 				###
###		This script relates to batch running using a GUI															###
###																													###
###		The script makes use of PowerFactory parallel processing and will require that the Parallel Processing		###
###		function in PowerFactory has been enabled and the number of cores has been set to N-1						###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###																													###
#######################################################################################################################

-------------------------------------------------------------------------------------------------------------------

"""
import time
import importlib
import sys
import pscharmonics

# Reload to allow repeat runs from PowerFactory and forcing constant resets
pscharmonics = importlib.reload(pscharmonics)

if __name__ == '__main__':
	"""
		Main function that is run
	"""
	# Initialise time counter for speed profiling
	t_start = time.time()

	# Initialise and run log message
	# logger = logging.getLogger(pscharmonics.constants.logger_name)
	# logger.setLevel(level=logging.DEBUG)
	sys.excepthook = pscharmonics.constants.logger.exception_handler
	logger = pscharmonics.constants.logger

	logger.info('Running {} in Graphical User Interface (GUI) Mode'.format(pscharmonics.constants.__title__))

	# Determine if running from PowerFactory and if so retrieve the current power factory version
	pf_version = pscharmonics.pf.running_in_powerfactory()

	main_gui = pscharmonics.gui.MainGui(pf_version=pf_version)

	# Capture final time and report complete
	t_end = time.time()
	logger.info(
		'GUI closed and any studies run after {:.0f} seconds'.format(t_end-t_start)
	)
