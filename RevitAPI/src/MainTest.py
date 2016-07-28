#===============================================================================
# Set up
#===============================================================================
# Standard:
from config.config import *

import logging.config
import unittest

from ExergyUtilities.utility_inspect import get_self
import utility_revit_api as util_ra
#===============================================================================
# Other modules
#===============================================================================
from win32com.client import Dispatch
#from Autodesk.Revit.DB import *
import ctypes as ctypes



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

import clr
clr.AddReference(r'C:\Program Files\Autodesk\Revit 2016\RevitAPI')
import Autodesk.Revit.DB as rvt_db
print(dir(rvt_db))

print(rvt_db)
#print(rvt_db.)


#clr.AddReference(r'C:\Program Files\Autodesk\Revit 2016\RevitAPIUI')

#from Autodesk.Revit.DB import *
#from Autodesk.Revit.DB.Architecture import *
#from Autodesk.Revit.DB.Analysis import *
#from Autodesk.Revit.UI.Selection import *
#from Autodesk.Revit.UI import *

print()

def main():
    logging.debug("Start")
    
    #import clr
    #import sys
    #import os
    #import re
    #import os.path as op
    #from datetime import datetime
    # import random as rnd
    # import pickle as pl
    # import time
    
    #clr.AddReference('PresentationCore')
    #clr.AddReference('RevitAPI')
    
    #raise
    #xl = Dispatch('Excel.Application')
    #rvt = Dispatch('RevitAPIUI')
#     print(xl)
#     print(ctypes.windll.kernel32.GetModuleHandleW(0))
#     path_dll = r"C:\Users\jon\AppData\Roaming\RevitPythonShell2016\CommandLoaderAssembly.dll"
#     lib = ctypes.WinDLL(path_dll)
#     print(lib)
#     print("Main")
    
if __name__ == '__main__':
    print("Test")
    main()
    pass