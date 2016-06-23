#===============================================================================
# pyRevit
#===============================================================================
__doc__ = 'Test'

#===============================================================================
# Logging
#===============================================================================
import logging
logging.basicConfig(level=logging.DEBUG)

#===============================================================================
# Imports Revit
#===============================================================================
try:
    uidoc = __revit__.ActiveUIDocument
    doc = __revit__.ActiveUIDocument.Document
    import Autodesk.Revit.DB as rvt_db
except:
    pass

#===============================================================================
# Imports other
#===============================================================================
import sys

path_package = r"C:\Dropbox\00 CAD Standards\70 Revit Python\Packages\xlrd"
sys.path.append(path_package)
import xlrd

#===============================================================================
# Main
#===============================================================================
logging.info("Python version : {}".format(sys.version))
try: 
    logging.info("uidoc : {}".format(uidoc))
    logging.info("doc : {}".format(doc))
except:
    pass

folder_excel_book = r"C:\CesCloud Revit\_03_IKEA_Working_Folder"
name_excel_book = r"\20160616 Document Register.xlsx"
path_excel_book = folder_excel_book + name_excel_book
logging.info("excel path : {}".format(path_excel_book))

xl_workbook = xlrd.open_workbook(path_excel_book)
logging.info("excel book : {}".format(xl_workbook))

xl_sheet = xl_workbook.sheet_by_index(0)
logging.info("excel sheet : {}".format(xl_sheet))