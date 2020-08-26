"""
#######################################################################################################################
###													batch_mode.py													###
###		Script deals with running the PSC harmonics studies in batch mode 											###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###																													###
#######################################################################################################################
"""

import os
import pscharmonics
import time

def run(pth_inputs=str(), test_settings=None, list_files=None):
	"""
		Function to run the study in batch mode with the inputs spreadsheet provided
	:param str pth_inputs:  Full path of settings file to import
	:param pscharmonics.file_io.StudyInputs test_settings:  Allows pre-loaded settings to be provided in cases where
															settings need to be adjusted slightly for testing
	:param list list_files:	List of files to run through for studies
	:return bool success:  Returns True if all studies run successfully
	"""
	# Initially success flag is set to False
	success = False

	# Get reference to the log message handler
	logger = pscharmonics.constants.logger

	# Determine if running from PowerFactory and if so retrieve the current power factory version
	pf_version = pscharmonics.pf.running_in_powerfactory()

	# Initialise PowerFactory instance
	pf = pscharmonics.pf.PowerFactory()
	pf.initialise_power_factory(pf_version=pf_version)

	# Make sure input works as list or single file
	multiple_studies=False
	if test_settings:
		list_files = [test_settings]
	elif pth_inputs and not list_files:
		list_files = [pth_inputs]
	elif list_files:
		multiple_studies=True
	else:
		raise ValueError('No inputs provided for running in batch mode')

	# Loop through to import settings or use test_settings
	for input_file in list_files:
		if multiple_studies:
			# Update UID
			pscharmonics.constants.uid = time.strftime('%Y%m%d_%H%M%S')

		if test_settings:
			inputs = input_file
		else:
			inputs = pscharmonics.file_io.StudyInputs(pth_file=input_file)



		# Create cases based on inputs file
		pf_projects = pscharmonics.pf.create_pf_project_instances(
			df_study_cases=inputs.cases,
			uid=pscharmonics.constants.uid,
			lf_settings=inputs.lf_settings,
			fs_settings=inputs.fs_settings
		)

		# Determine whether to run and export a pre-case check
		if inputs.settings.pre_case_check:
			pre_case_check_file = os.path.join(
				inputs.settings.export_folder,
				'Pre Case Check_{}.xlsx'.format(pscharmonics.constants.uid)
			)

			# Run the pre-case check
			pscharmonics.pf.run_pre_case_checks(
				pf_projects=pf_projects,
				terminals=inputs.terminals,
				include_mutual=inputs.settings.export_mutual,
				export_pth=pre_case_check_file,
				contingencies=inputs.contingencies,
				contingencies_cmd=inputs.contingency_cmd
			)

		# Update results folder to include the results file_name
		pth_results = os.path.join(inputs.settings.export_folder, inputs.settings.results_name)
		inputs.settings.add_folder(pth_results_file=pth_results)

		# Run the full study
		# Iterate through each project and create the various cases, the includes running a pre-case check but no
		# output is saved at this point
		_ = pscharmonics.pf.run_studies(pf_projects=pf_projects, inputs=inputs)


		# Determine whether results should be exported to excel
		if inputs.settings.export_to_excel:
			# Export results to the path detailed in the inputs spreadsheet
			_ = pscharmonics.file_io.ExtractResults(target_file=pth_results, search_paths=(inputs.settings.export_folder,))

			# Confirm the file exists to set as a status flag
			if os.path.isfile(pth_results):
				success = True
		else:
			logger.info(
				(
					'The inputs requested that the results were not exported to excel and therefore no results have been '
					'produced.  If you require results you will either need to change the input setting or use the GUI to '
					'combine the results that have been saved in the folder:\n\t{}'
				).format(inputs.settings.export_folder)
			)
			success = True

		# Determine whether temporary power factory folders should be deleted
		if inputs.settings.delete_created_folders:
			for prj_name, prj in pf_projects.items():  # type: str, pscharmonics.pf.PFProject
				prj.delete_temp_folders()

	return success

