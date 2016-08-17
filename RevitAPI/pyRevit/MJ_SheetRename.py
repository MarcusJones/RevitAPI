from __future__ import print_function

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
# Configuration
#===============================================================================
#Overview plan, center
OVERVIEWPLAN_CENTER = rvt_db.XYZ(1.656061764, 1.323727996, 0.036311942)
LOCATIONPLAN_CENTER = rvt_db.XYZ(3.539872609, 0.685768315, 0.000000000)
LEGEND_CENTER = rvt_db.XYZ(3.522094284, 1.583534318, 0.000000000)

#===============================================================================
# Imports other
#===============================================================================
import sys
import csv

path_package = r"C:\EclipseGit\ExergyUtilities\ExergyUtilities"
sys.path.append(path_package)
import utility_revit_api as util_ra

#===============================================================================
# Definitions
#===============================================================================
def rename_sheets(data_dict, sheets_by_name):
    logging.debug(util_ra.get_self())
    cnt_updated = 0
    for i,row in enumerate(data_dict):
        assert row['SOURCE'] == 'RVT', "Only works with RVT drawings, not [{}]".format(row['SOURCE'])
        
        #if row['OLD NAME']:
        #    assert row['OLD NAME'] in sheets_by_name, "{} sheet not found in project".format(row['OLD NAME'])
        #else:
        #    continue
        
        # Select this sheet to update
        this_sheet = sheets_by_name[row['Sheet Name']]
        
        # Change OLD NAME to Sheet Name column
#         util_ra.change_parameter(rvt_doc, 
#                                  this_sheet, 
#                                  'Sheet Name', 
#                                  row['Sheet Name'])
        
        # Update Sheet Number from Sheet Number column
        try:
            util_ra.change_parameter(rvt_doc, 
                                     this_sheet, 
                                     'Sheet Number', 
                                     row['Sheet Number'])
            cnt_updated += 1     
        except: 
            logging.error("Couldn't update parameter!".format())
            
        logging.debug("Updated {}".format(row['Sheet Name']))
    
    logging.debug("Successfully change {} out of {} sheets".format(cnt_updated,i))
    

#===============================================================================
# Main
#===============================================================================
#-Logging info---
logging.info("Python version : {}".format(sys.version))
logging.info("uidoc : {}".format(rvt_uidoc))
logging.info("doc : {}".format(rvt_doc))
logging.info("app : {}".format(rvt_app))

#-Paths---
folder_csv = r"C:\CesCloud Revit\_03_IKEA_Working_Folder"
name_csv = r"\20160728 Document Register.csv"
path_csv = folder_csv + name_csv

#-Get data---
data_dict = util_ra.get_data_csv(path_csv)
data_dict_RVT = [row for row in data_dict if row['SOURCE'] == 'RVT']

#-Get all floorplans, sheets_by_name, titleblocks, legends---
util_ra.get_all_views(rvt_doc)
title_blocks = util_ra.get_title_blocks(rvt_doc)
sheets_by_name = util_ra.get_sheet_dict_by_names(rvt_doc)
floorplans = util_ra.get_views_by_type(rvt_doc, 'FloorPlan')
all_views = util_ra.get_all_views(rvt_doc)
legends = util_ra.get_views_by_type(rvt_doc,'Legend')
templates = util_ra.get_view_templates(rvt_doc)

#-Rename---
rename_sheets(data_dict_RVT, sheets_by_name)

logging.info("---DONE---".format())

