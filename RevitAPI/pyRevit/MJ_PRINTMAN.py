from __future__ import print_function

#===============================================================================
# pyRevit
#===============================================================================
__doc__ = """

"""

#===============================================================================
# Logging
#===============================================================================
import logging
logging.basicConfig(level=logging.DEBUG)

#===============================================================================
# Import Revit
#===============================================================================
rvt_app = __revit__.Application
rvt_uidoc = __revit__.ActiveUIDocument
rvt_doc = __revit__.ActiveUIDocument.Document
import Autodesk.Revit.DB as rvt_db

#===============================================================================
# Configuration
#===============================================================================

#===============================================================================
# Imports other
#===============================================================================
import sys
import csv
import time

path_package = r"C:\EclipseGit\ExergyUtilities\ExergyUtilities"
sys.path.append(path_package)
import utility_revit_api as util_ra

#===============================================================================
# Definitions
#===============================================================================
def PrintView(doc, sheet, pRange, printerName, combined, filePath, printSetting):
    # create view set for this one sheet
    viewSet = rvt_db.ViewSet()
    viewSet.Insert(sheet)
    
    # determine print range
    printManager = doc.PrintManager
    printManager.PrintRange = pRange
    printManager.Apply()
    
    # make new view set current
    viewSheetSetting = printManager.ViewSheetSetting
    viewSheetSetting.CurrentViewSheetSet.Views = viewSet
    
    # set printer
    printManager.SelectNewPrintDriver(printerName)
    printManager.Apply()
    
    # set combined and print to file
    if printManager.IsVirtual:
        printManager.CombinedFile = combined
        printManager.Apply()
        printManager.PrintToFile = True
        printManager.Apply()
    else:
        printManager.CombinedFile = combined
        printManager.Apply()
        printManager.PrintToFile = False
        printManager.Apply()
        
    # set file path
    printManager.PrintToFileName = filePath
    printManager.Apply()
    
    # apply print setting
    try:
        printSetup = printManager.PrintSetup
        printSetup.CurrentPrintSetting = printSetting
        printManager.Apply()
    except:
        pass
    
    # save settings and submit print
    rvt_db.TransactionManager.Instance.EnsureInTransaction(doc)
    viewSheetSetting.SaveAs("tempSetName")
    printManager.Apply()
    printManager.SubmitPrint()
    viewSheetSetting.Delete()
    rvt_db.TransactionManager.Instance.TransactionTaskDone()
    
    return True


#===============================================================================
# Main
#===============================================================================
#-Logging info---
logging.info("Python version : {}".format(sys.version))
logging.info("uidoc : {}".format(rvt_uidoc))
logging.info("doc : {}".format(rvt_doc))
logging.info("app : {}".format(rvt_app))

viewsets = rvt_db.FilteredElementCollector(rvt_doc).OfClass(rvt_db.ViewSheetSet)

for vs in viewsets:
    print(vs)
    #for sheet in vs:
    #    print(vs.Name)
    
print_man = rvt_doc.PrintManager

print(print_man)

logging.info("---DONE---".format())

