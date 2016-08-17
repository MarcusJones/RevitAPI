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
# Import Revit
#===============================================================================
rvt_app = __revit__.Application
rvt_uidoc = __revit__.ActiveUIDocument
rvt_doc = __revit__.ActiveUIDocument.Document
import Autodesk.Revit.DB as rvt_db

#===============================================================================
# Import Excel
#===============================================================================
#import clr
#clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
#from Microsoft.Office.Interop import Excel

#===============================================================================
# Imports other
#===============================================================================
import sys

#path_package = r"C:\Dropbox\00 CAD Standards\70 Revit Python\Packages\xlrd"
#sys.path.append(path_package)
#import xlrd

path_package = r"C:\EclipseGit\ExergyUtilities\ExergyUtilities"
sys.path.append(path_package)
import utility_excel_api_Interop as util_excel
import utility_revit_api as util_ra

#===============================================================================
# Definitions
#===============================================================================

#===============================================================================
# Main
#===============================================================================
#-Logging info---
logging.info("Python version : {}".format(sys.version))
logging.info("uidoc : {}".format(rvt_uidoc))
logging.info("doc : {}".format(rvt_doc))
logging.info("app : {}".format(rvt_app))

#-Get all floorplans, sheets, titleblocks, legends---
# util_ra.get_all_views(rvt_doc)
# title_blocks = util_ra.get_title_blocks(rvt_doc)
# sheets = util_ra.get_sheet_dict(rvt_doc)
# floorplans = util_ra.get_views_by_type(rvt_doc, 'FloorPlan')
# legends = util_ra.get_views_by_type(rvt_doc,'Legend')
# templates = util_ra.get_view_templates(rvt_doc)

#util_ra.selection(rvt_uidoc,rvt_doc)

#util_ra.inspect_selection(this_el)

#util_ra.table_parameters(this_el)

#util_ra.document_parameters(rvt_doc)
#this_el = util_ra.single_selection(rvt_uidoc,rvt_doc)
#util_ra.list_parameters(this_el)

this_view = rvt_doc.ActiveView

#util_ra.table_parameters(this_view)
util_ra.change_parameter(rvt_doc, this_view, "SHEETNAME_ORIGINATOR", "5MEC NEW")

#this_el.get_parameter()
#util_ra.project_parameters(rvt_doc)



#for item in dir(this_el):
#    print(item)

#single_selection
#util_ra.document_parameters(rvt_doc)
#util_ra.get_element_by_id(doc,id)

logging.info("---DONE---".format())


#xl_sheet = xl_workbook.sheet_by_index(0)
#logging.info("excel sheet : {}".format(xl_sheet))