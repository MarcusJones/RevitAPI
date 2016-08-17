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
# Configuration
#===============================================================================
OVERVIEWPLAN_CENTER = rvt_db.XYZ(1.656061764, 1.323727996, 0.036311942)

#===============================================================================
# Imports other
#===============================================================================
import sys
from collections import defaultdict


path_package = r"C:\EclipseGit\ExergyUtilities\ExergyUtilities"
sys.path.append(path_package)
import utility_revit_api as util_ra

#===============================================================================
# Definitions
#===============================================================================
def update_sheets_views(excel_dict, sheets_by_name, views, viewports):
    logging.debug(util_ra.get_self())
    count = 0
    for i,row in enumerate(excel_dict):
        
        assert row['SOURCE'] == 'RVT'
        if row['VIEW TYPE'] != 'PLAN':
            continue

        assert row['Sheet Name'] in sheets_by_name, "Sheet {} does not exist".format(row['Sheet Name'])
        assert row['View Name'] in views, "View {} does not exist".format(row['View Name'])
        assert row['MAIN VIEWPORT'] in viewports, "Viewport {} does not exist".format(row['MAIN VIEWPORT'])

        # Test views in register matches views placed in Revit
        placed_views = sheets_by_name[row['Sheet Name']].GetAllPlacedViews() 
        view_match = False
        legend_match = False
        for view_id in placed_views:
            view = rvt_doc.GetElement(view_id)
            if view.Name == row['View Name']:
                view_match = True
                
        # Test view match
        if not view_match:
            count+=1
            print("Add {} to {}".format(row['View Name'],row['Sheet Name']))
            util_ra.add_view_sheet(rvt_doc, 
                                   sheets_by_name[row['Sheet Name']], 
                                   views[row['View Name']], 
                                   OVERVIEWPLAN_CENTER,
                                   viewports[row['MAIN VIEWPORT']].GetTypeId())
                
    logging.debug("Added views to {} sheets".format(count))
                

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
name_csv = r"\20160722 Document Register.csv"
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
viewports = util_ra.get_viewports_dict_by_names(rvt_doc)

#-Check---
update_sheets_views(data_dict_RVT,sheets_by_name,all_views,viewports)

logging.info("---DONE---".format())