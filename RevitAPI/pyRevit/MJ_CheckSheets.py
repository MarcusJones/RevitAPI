from __future__ import print_function

#===============================================================================
# pyRevit
#===============================================================================
__doc__ = 'Script docstring'

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
# Imports other
#===============================================================================
import sys
from collections import defaultdict

path_package = r"C:\EclipseGit\ExergyUtilities\RevitUtilities"
sys.path.append(path_package)
import utility_revit_api2 as util_ra
import utility as util

#===============================================================================
# Definitions
#===============================================================================
def check_sheets(excel_dict, sheets_by_name, views, view_templates):
    logging.debug(util.get_self())
    
    counts = defaultdict(int)
    
    for i,row in enumerate(excel_dict):
        flg_error = False
        counts["Total Rows"] +=1
        assert row['SOURCE'] == 'RVT'
        
        counts["Total Rows in RVT"] += 1
        
        logging.info("CHECK: {} {}".format(row['Sheet Name'],row['Sheet Number'],))

        # Test sheet exists
        if row['Sheet Name'] not in sheets_by_name:
            counts["Missing Sheet"]+=1
            logging.error("Sheet {} does not exist".format(row['Sheet Name']))
            flg_error = True
            
            
        # Only PLAN VIEWS
        if row['VIEW TYPE'] == 'PLAN':            
            # Test view exists
            if row['View Name'] not in views:
                counts["Missing View"] +=1
                logging.error("View {} does not exist".format(row['View Name']))
                flg_error = True

            placed_views = sheets_by_name[row['Sheet Name']].GetAllPlacedViews() 
            view_match = False
            legend_match = False
            for view_id in placed_views:
                view = rvt_doc.GetElement(view_id)
                if view.Name == row['View Name']:
                    view_match = True
                if view.Name == row['LOCATION LEGEND']:
                    legend_match = True
                  
            # Test view match
            if not view_match:
                counts["Missing or incorrect view in Revit"]+=1
                logging.error("View {} not placed on {}".format(row['View Name'], row['Sheet Name']))
                flg_error = True
                

            # Test location legend match
            if not legend_match:
                counts["Missing or incorrect location part in Revit"]+=1
                logging.error("Location legend {} not placed on {}".format(row['Sheet Name'], row['LOCATION LEGEND']))
                flg_error = True
        
        if not flg_error:
            logging.info("*** OK ***".format())
        else:
            logging.error("XXX XXX ERROR XXX XXX".format())

    for k in counts:
        print("{}:{}".format(k, counts[k]))

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
name_csv = r"\20160729 Document Register.csv"
path_csv = folder_csv + name_csv

#-Get data---
data_dict = util.get_data_csv(path_csv)
data_dict_RVT = [row for row in data_dict if row['SOURCE'] == 'RVT']

#-Get all floorplans, sheets_by_name, titleblocks, legends---
util_ra.get_all_views(rvt_doc)
title_blocks = util_ra.get_title_blocks(rvt_doc)
sheets_by_name = util_ra.get_sheet_dict_by_names(rvt_doc)
floorplans = util_ra.get_views_by_type(rvt_doc, 'FloorPlan')
all_views = util_ra.get_all_views(rvt_doc)
legends = util_ra.get_views_by_type(rvt_doc,'Legend')
templates = util_ra.get_view_templates(rvt_doc)

#-Check---
check_sheets(data_dict_RVT,sheets_by_name,all_views,templates)

logging.info("---DONE---".format())