#===============================================================================
# Set up
#===============================================================================
# Standard:
from config.config import *

import logging.config
#import unittest


#===============================================================================
# Logging
#===============================================================================
print(ABSOLUTE_LOGGING_PATH)
logging.config.fileConfig(ABSOLUTE_LOGGING_PATH)
myLogger = logging.getLogger()
myLogger.setLevel("DEBUG")

#===============================================================================
# Utilities
#===============================================================================
#import RevitUtilities.utility_revit as util_ra
import RevitUtilities.utility_get_elements as util_get_el


#===============================================================================
# Get System (???)
#===============================================================================

import clr
clr.AddReference('System')                 # Enum, Diagnostics
clr.AddReference('System.Collections')     # List

# Core Imports
from System import Enum
from System.Collections.Generic import List
from System.Diagnostics import Process

#===============================================================================
# Get Revit
#===============================================================================


import sys
sys.path.append(r'C:\Program Files\Autodesk\Revit 2017')
#sys.path.append(r'C:\Program Files\Autodesk\Revit 2017')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUIMacros')


import clr

# Revit namespace
clr.AddReference(r'C:\Program Files\Autodesk\Revit 2017\RevitAPI')

# Get DB
import Autodesk.Revit.DB as rvt_db
rvt_doc = rvt_db.Document
logging.debug("Added {}".format(rvt_db))
logging.debug("Added {}".format(rvt_doc))
from Autodesk.Revit.DB import FilteredElementCollector


clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import * 
 
app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document


#from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory

#raise
#print(rvt_db.Document)
# Get UI
import Autodesk.Revit.UI as rvt_ui
logging.debug("Added {}".format(rvt_ui))

# Revit UI namespace
#clr.AddReference(r'C:\Program Files\Autodesk\Revit 2017\RevitAPIUI')
#rvt_uidoc = Autodesk.Revit.ActiveUIDocument
#rvt_doc = rvt_db.ActiveUIDocument.Document

#raise

#Dim uidoc As UIDocument = commandData.Application.ActiveUIDocument

# DOES NOT WORK


# clr.AddReference('RevitAPIUI')


#doc = rvt_db.doc
#import Autodesk.Revit.UI.UIDocument as rvt_ui_doc
#print(doc)


#===============================================================================
# Main script
#===============================================================================

#import clr
#clr.AddReference(r'C:\Program Files\Autodesk\Revit 2016\RevitAPI')
#import Autodesk.Revit.DB as rvt_db
#print(dir(rvt_db))

#util_get_el.get_element_by_id()
print('test')
walls = util_get_el.get_element_OST_Walls_Document(rvt_doc,rvt_db,FilteredElementCollector)
print(walls)
#el_dict_instances = util_ra.get_sort_all_FamilyInstance(rvt_doc)
raise 

print(rvt_db)
print(rvt_ui)
print(rvt_ui_doc)

"""
using System;
using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
"""

print(rvt_db.__doc__)
print(rvt_db.__name__)
print(rvt_db.__file__)
print(rvt_db.Document)
raise
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = list(__revit__.ActiveUIDocument.Selection.Elements)


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