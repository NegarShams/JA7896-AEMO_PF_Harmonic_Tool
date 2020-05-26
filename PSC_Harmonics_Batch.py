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
import logging
import time
import pscharmonics

input_spreadsheet_name = 'PSC_Harmonics_Inputs.xlsx'

if __name__ == '__main__':
	"""
		Main function that is run
	"""
	# Initialise time counter for speed profiling
	t_start = time.time()

	# Initialise and run log message
	logger = logging.getLogger(pscharmonics.constants.logger_name)
	logger.info('Batch Study Run using Input Filename: {}'.format(input_spreadsheet_name))

	# Retrieve inputs
	pth_inputs = os.path.join(os.path.dirname(__file__), input_spreadsheet_name)
	inputs = pscharmonics.file_io.StudyInputsDev(pth_file=pth_inputs)

	# Initialise PowerFactory instance
	pf = pscharmonics.pf.PowerFactory()
	pf.initialise_power_factory()


	# Create cases based on inputs file
	pf_projects = pscharmonics.pf.create_pf_project_instances(
		df_study_cases=inputs.cases,
		uid=pscharmonics.constants.uid,
		lf_settings=inputs.lf_settings,
		fs_settings=inputs.fs_settings,
		export_pth=inputs.settings.export_folder
	)

	# Iterate through each project and create the various cases, the includes running a pre-case check but no
	# output is saved at this point
	for project_name, project in pf_projects.items():
		project.create_cases(
			study_settings=inputs.settings,
			export_pth=inputs.settings.export_folder,
			contingencies=inputs.contingencies,
			contingencies_cmd=inputs.contingency_cmd
		)

		# Update the auto executable for this project
		project.update_auto_exec()

		# Batch run the results
		project.task_auto.Execute()

		# Delete temporary folders created for this project
		project.delete_temp_folders()

	# Capture final time and report complete
	t_end = time.time()
	logger.info(
		'Study completed in {:.0f} seconds with results saved in folder {}'.format(
			t_end-t_start, inputs.settings.export_folder
		)
	)
