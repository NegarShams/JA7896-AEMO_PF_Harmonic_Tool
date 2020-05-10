"""
#######################################################################################################################
###													pf																###
###		Script deals with writing data to PowerFactory and any processing that takes place which requires 			###
###		interacting with power factory																				###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###		project JI6973 for EirGrid project PSPF010 - Specialise Support in Power Quality Analysis during 2018		###
###																													###
#######################################################################################################################
"""

import os
import sys
import math
import pscharmonics.constants as constants
import multiprocessing
import time
import logging
import distutils.version

# powerfactory will be defined after initialisation by the PowerFactory class
powerfactory = None
app = None

# Meta Data
__author__ = 'David Mills'
__version__ = '2.1.2'
__email__ = 'david.mills@pscconsulting.com'
__phone__ = '+44 7899 984158'
__status__ = 'In Development - Beta'


# TODO: Loop through DataFrame of study cases by project and create project references
# TODO: Create testing routine for