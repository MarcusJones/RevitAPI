from __future__ import print_function

#===============================================================================
# pyRevit
#===============================================================================
__doc__ = """
Deletes viewport

Creates new viewport with specified type, location, and view

Easier to Delete and replace!
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
def update_sheets_views_viewport(excel_dict, sheets_by_name, views, view_ports):
    logging.debug(util_ra.get_self())
    count = 0
    count_added = 0

    for i,row in enumerate(excel_dict):
        count += 1
        assert row['SOURCE'] == 'RVT'
        if row['VIEW TYPE'] != 'PLAN':
            continue
        assert row['Sheet Name'] in sheets_by_name, "Sheet {} does not exist".format(row['Sheet Name'])
        assert row['View Name'] in views, "View {} does not exist".format(row['View Name'])
        assert row['MAIN VIEWPORT'] in view_ports, "Viewport {} does not exist".format(row['MAIN VIEWPORT'])
        
        #logging.debug("Processing {} with {}".format(row['Sheet Name'],row['View Name']))
        
#         placed_views = sheets_by_name[row['Sheet Name']].GetAllPlacedViews() 
#         for view_id in placed_views:
#             
#             view = rvt_doc.GetElement(view_id)
#             #print("\t{} placed on sheet".format(view))
#             
#             #view_port.ChangeTypeId(viewport_ID)
#             #if view.Name == row['View Name']:
#             #    view_match = True
#             
        for viewport_id in sheets_by_name[row['Sheet Name']].GetAllViewports():
            #print(viewport_id)
            
            viewport = rvt_doc.GetElement(viewport_id)
            view_id = viewport.ViewId
            this_view = rvt_doc.GetElement(view_id)
            
            #print(viewport, this_view)
            #print(viewport, this_view.Name)
            flg_found = False
            if this_view.Name == row['View Name']:
                logging.debug("View {} found on sheet {}".format(this_view.Name,row['Sheet Name']))
                #print(viewport.Location.ToString())
                #util_ra.print_dir(viewport.Location)
                #raise
                flg_found = True
                break 
            
        #assert flg_found, "{} not found on sheet {}".format(row['View Name'],row['Sheet Name'])
        
        # Delete viewport
        with util_ra.Trans(rvt_doc,"Delete viewport"):
            try:
                rvt_doc.Delete(viewport.Id)
            except:
                logging.error("Couldn't delete viewport".format())
                
        # Replace viewport
        try:
            util_ra.add_view_sheet(rvt_doc, 
                                   sheets_by_name[row['Sheet Name']], 
                                   views[row['View Name']], 
                                   OVERVIEWPLAN_CENTER,
                                   view_ports[row['MAIN VIEWPORT']].GetTypeId()
                                   )
            count_added += 1
        except:
            logging.error("Couldn't recreate viewport with ID {}".format(view_ports[row['MAIN VIEWPORT']].GetTypeId()))
            
    print("Added {} viewports over {} rows")

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
update_sheets_views_viewport(data_dict_RVT,sheets_by_name,all_views,viewports)

logging.info("---DONE---".format())