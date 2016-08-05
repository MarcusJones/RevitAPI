#===============================================================================
# Set up
#===============================================================================
# Standard:
from config.config import *
import re
import logging.config
import unittest

from ExergyUtilities.utility_inspect import get_self
#from sympy.categories.baseclasses import Category
#from statsmodels.genmod.families.family import Family

from collections import defaultdict
import sys

import csv
#===============================================================================
# Other modules
#===============================================================================

#===============================================================================
# Logging
#===============================================================================
print(ABSOLUTE_LOGGING_PATH)
logging.config.fileConfig(ABSOLUTE_LOGGING_PATH)
myLogger = logging.getLogger()
myLogger.setLevel("DEBUG")

#===============================================================================
# Main script
#===============================================================================
logging.info("Python version : {}".format(sys.version))
#logging.info("uidoc : {}".format(rvt_uidoc))
#logging.info("doc : {}".format(rvt_doc))
#logging.info("app : {}".format(rvt_app))
 

def main():
    logging.debug("Start")
    
    logging.debug("Done")


if __name__ == '__main__':
    main()
