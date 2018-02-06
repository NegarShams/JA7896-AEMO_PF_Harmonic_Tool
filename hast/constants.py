"""
	Script to contain all the custom constants (could use an .init file instead)
"""
import os
import sys

__script_name__ = os.path.basename(__file__)
pf_version = '15.2'

class PowerFactory:
	"""
		Class will contain all the paths that relate to the PowerFactory software and are dependant on the version
		that is being used
	"""
	def __init__(self, version):
		"""
			Initialise class
		:param str version: Version number specifying the version of powerfactory that is being run
		"""
		if version == '15.2':
			self.pf_path = r'C:\Program Files\DIgSILENT\PowerFactory 15.2'
			self.version = version
		elif version == '16.0.5':
			self.pf_path = r'C:\Program Files\DIgSILENT\PowerFactory 2016 SP5'
			self.version = version
		elif version == '17.0.6':
			self.pf_path = r'C:\Program Files\DIgSILENT\PowerFactory 2017 SP6'
			self.version = version
		else:
			# Error message since python version is not known
			raise SyntaxError('Directory for Python version {} is not defined, please update {}.{}'
							  .format(version, __script_name__, self.__class__.__name__))

		self.pf_python_path = self.get_python_path()
	def get_python_path(self):
		"""
			Determines the path to the python files within the PowerFactory folder based on the version of
			python that is running
		:return: str python_path:  Full path to folder which contains powerfactory.pyd
		"""
		# Get running python version
		py_version = sys.version_info
		py_version = int(py_version[0]) + int(py_version[1]) / 10.0
		python_folder = os.path.join(self.pf_path, 'Python', '{}'.format(py_version))


		# In case python version does not exist then decrease the python version number
		if not os.path.isdir(python_folder):
			raise SyntaxError(('Power Factory folder {} cannot be found for version {} of python, ' +
							  'please check you are running the correct version')
							  .format(python_folder, py_version))

		return python_folder



	