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
def update_sheet_parameters(excel_dict, sheets_by_name, flg_update=False, verbose=False):
    logging.debug(util_ra.get_self())
    count_total_rows = 0
    count_total_rvt = 0
    focus_params =[
                   #'Project Number', # PROJECT PARAM
                   #'Project Status', # PROJECT PARAM
                   'Designing Company', # YES
                   'Project Manager MEP', # YES
                   'Project Manager ID',# YES
                   'Designer', # YES
                   'Designer ID', # YES
                   'Draftsperson', # YES
                   #'Sheet Name' # PROJECT PARAM
                   'Design', # YES
                   'Design No.', # YES
                   #'Drawing Type',
                   #'Project Status',
                   'Sheet Issue Date',
                   'Revision',
                   ]
    for i,row in enumerate(excel_dict):
        count_total_rows += 1
        if row['SOURCE'] == 'RVT':
            count_total_rvt+=1
            # Get sheet (exists)
            try: this_sheet = sheets_by_name[row['Sheet Name']]
            except:
                logging.error("Sheet {} does not exist".format(sheet_name))
                raise
            
            sheet_name = row['Sheet Name']
            
            for k in row:
                if k in focus_params:
                    # 
                    flg_exists = util_ra.parameter_exists(this_sheet, k)
                    if flg_exists:
                        value = util_ra.get_parameter_value(this_sheet, k)
                    else:
                        value = ""
                    
                    if row[k] == value:
                        flg_match = True
                    else:
                        flg_match = False
                        
                    if verbose:
                        print("Row {0} {1:20} Exists:{2:10} Value:{3:20} New Value:{4:20} Match:{5:10}".format())
                              
                              
                else:
                    if verbose:
                        print("Skipped {}".format(k))
                
                if k in focus_params and not flg_match:
                    util_ra.change_parameter(rvt_doc, this_sheet, k, row[k])
            #util_ra.table_parameters(this_sheet)
            
    logging.debug("Iterated {} document rows".format(count_total_rows))
    logging.debug("Checked {} sheets_by_name (RVT plan views only)".format(count_total_rvt))

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
data_dict = get_data_csv(path_csv)

# data_dict_RVT = list()
# for row in data_dict:
#     for i,row in enumerate(excel_dict):
#         count_total_rows += 1
#         if row['SOURCE'] == 'RVT' and row['VIEW TYPE'] == 'PLAN':
#             data_dict_RVT.append()
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
update_sheet_parameters(data_dict_RVT,sheets_by_name,flg_update=False)

logging.info("---DONE---".format())