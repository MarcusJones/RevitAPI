#===============================================================================
# Set up
#===============================================================================
# Standard:
from config.config import *
import re
import logging.config
import unittest

import os

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

def main():
    logging.debug("Start")
    
    folder = r"C:\Users\jon\Desktop\PDF Print"
    
    for i,filename in enumerate(os.listdir(folder)):
        #if filename.startswith("cheese_"):
        #    os.rename(filename, filename[7:])
        new_name = filename.replace("C_Users_jon_Documents_160315_IKEA_MEP_Central_jon - Sheet - ","")
        new_name = new_name.replace("--","")

        print(new_name)
        print(i,filename)
        
    logging.debug("Done")


if __name__ == '__main__':
    main()
