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
def update_view_parameters(excel_dict, all_views):
    logging.debug(util_ra.get_self())
    count_total_rows = 0
    count_total_rvt = 0

    for i,row in enumerate(excel_dict):
        count_total_rows += 1

        assert row['SOURCE'] == 'RVT'
        if row['VIEW TYPE'] != 'PLAN':
            continue
        assert row['View Name'] in all_views, "View {} does not exist".format(row['View Name'])
        # Get view (exists)
       
        logging.debug("Processing {} rows".format(row['View Name']))
        
        this_view = all_views[row['View Name']]
        
        # Update scale
        #flg_exists = util_ra.parameter_exists(this_view, 'Scale Value    1:')
        #util_ra.table_parameters(this_view)
        #print(this_view)
        logging.debug("Changing scale of view from {} to {}".format(this_view.Scale,row['SCALE']))
        
        with util_ra.Trans(rvt_doc,"Change scale"):
            this_view.Scale = int(row['SCALE'])
        
    #logging.debug("Iterated {} document rows".format(count_total_rows))
    #logging.debug("Checked {} sheets_by_name (RVT plan views only)".format(count_total_rvt))

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
name_csv = r"\20160721 Document Register.csv"
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

#-Run---
update_view_parameters(data_dict_RVT,all_views)

logging.info("---DONE---".format())