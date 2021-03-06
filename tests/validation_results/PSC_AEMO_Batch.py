"""
#######################################################################################################################
###											PSC Harmonics															###
###		Script produced by David Mills (PSC) for Automated Running of Frequency Scans in PowerFactory 				###
###																									 				###
###		This script relates to batch running using an input spreadsheet rather than running via a GUI 				###
###																													###
###		The script makes use of PowerFactory parallel processing and will require that the Parallel Processing		###
###		function in PowerFactory has been enabled and the number of cores has been set to N-1						###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###																													###
#######################################################################################################################

-------------------------------------------------------------------------------------------------------------------

"""
import os
import time
import pscharmonics

input_spreadsheet_names = (
	'Validation_Inputs.xlsx',
	'Validation_Inputs_10.xlsx',
	'Validation_Inputs_100.xlsx',
	'Validation_Inputs_Loci10.xlsx',
	'Validation_Inputs_Loci10_includeall.xlsx',
	'Validation_Inputs_Loci10_unlim.xlsx',
	'Validation_Inputs_Loci100.xlsx',
	'Validation_Inputs_Loci100mutual.xlsx'
)

# Force to use latest PowerFactory version
pscharmonics.constants.PowerFactory.target_power_factory = str()

if __name__ == '__main__':
	"""
		Main function that is run
	"""
	# Initialise time counter for speed profiling
	t_start = time.time()

	# Initialise and run log message
	# logger = logging.getLogger(pscharmonics.constants.logger_name)
	logger = pscharmonics.constants.logger
	logger.info('Batch Study Run using Input Filenames: \n\t{}'.format(
		'\n\t'.join(input_spreadsheet_names)
	))

	# Establish inputs file name
	pth_inputs = [os.path.join(os.path.dirname(__file__), name) for name in input_spreadsheet_names]

	# Run batch study
	success = pscharmonics.batch_mode.run(list_files=pth_inputs)

	# Capture final time and report complete
	t_end = time.time()
	if success:
		logger.info(
			'Study completed in {:.0f} seconds successfully'.format(t_end-t_start)
		)
	else:
		logger.critical(
			(
				'An error has occurred and after {:.0f} seconds results have not been produced as expected.  Check the '
				'messages displayed above to determine the issue'
			).format(t_end-t_start)
		)
