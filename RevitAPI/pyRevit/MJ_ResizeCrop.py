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
#Overview plan, center
OVERVIEWPLAN_CENTER = rvt_db.XYZ(1.656061764, 1.323727996, 0.036311942)
LOCATIONPLAN_CENTER = rvt_db.XYZ(3.539872609, 0.685768315, 0.000000000)
LEGEND_CENTER = rvt_db.XYZ(3.522094284, 1.583534318, 0.000000000)


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
def resize_crop(excel_dict, sheets_by_name,all_views):
    
    logging.debug(util_ra.get_self())
    count = 0
    for i,row in enumerate(excel_dict):
        assert row['SOURCE'] == 'RVT'
        if row['VIEW TYPE'] != 'PLAN':
            continue
        if row['PART'] not in ['1','2','3','4','5','6']:
            continue
        assert row['Sheet Name'] in sheets_by_name, "Sheet {} does not exist".format(row['Sheet Name'])
        assert row['View Name'] in all_views, "View {} does not exist".format(row['View Name'])
        
        
        #assert row['LOCATION LEGEND'] in legends, "View {} does not exist".format(row['View Name'])
        #assert row['LOC LEG VIEWPORT'] in all_viewports, "Viewport family {} does not exist".format(row['LOC LEG VIEWPORT'])
        
        print(row['Sheet Name'],row['View Name'])
        
        this_view = all_views[row['View Name']]
        print(this_view.CropBox)
        raise
#         # Test views in register matches views placed in Revit
#         placed_views = sheets_by_name[row['Sheet Name']].GetAllPlacedViews() 
#         view_match = False
#         legend_match = False
#         for view_id in placed_views:
#             view = rvt_doc.GetElement(view_id)
#             if view.Name == row['LOCATION LEGEND']:
#                 view_match = True
#         
#         # Test view match
#         if not view_match:
#             count+=1
#             new_viewport_type = all_viewports[row['LOC LEG VIEWPORT']]
#             new_vp_id = new_viewport_type.GetTypeId()
#                         
#             logging.debug("Adding {} to {} using {}".format(row['LOCATION LEGEND'],row['Sheet Name'],new_viewport_type.Name))
#             view_port = util_ra.add_view_sheet(rvt_doc, 
#                                                sheets_by_name[row['Sheet Name']], 
#                                                legends[row['LOCATION LEGEND']], 
#                                                LOCATIONPLAN_CENTER,
#                                                new_vp_id)
#             
            
    #logging.debug("Added location views to {} sheets".format(count))
                

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
name_csv = r"\20160712 Document Register.csv"
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
all_viewports = util_ra.get_viewports_dict_by_names(rvt_doc)

#-Run---

resize_crop(data_dict_RVT,sheets_by_name,all_views)

logging.info("---DONE---".format())