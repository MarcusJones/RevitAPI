from __future__ import print_function

#===============================================================================
# pyRevit
#===============================================================================
__doc__ = """
Collect data from CSV table, headers in second row
Apply parameters to sheet objects from table

3.OCT.2016 - Updated for OD phase
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
import time

#path_package = r"C:\EclipseGit\ExergyUtilities\RevitUtilities"
path_package = r"C:\EclipseGit\ExergyUtilities"
sys.path.append(path_package)
import RevitUtilities as util_ra
#import utility_general as util_gen

#===============================================================================
# Definitions
#===============================================================================
def update_sheet_parameters(excel_dict, sheets_by_name, flg_update=False, verbose=False):
    logging.debug(util_ra.get_self())
    count_total_rows = 0
    count_total_rvt = 0
    count_total_updated = 0
    count_total_failed = 0
    failed_sheets = list()
    focus_params =[
                   #'Project Number', # PROJECT PARAM
                   #'Project Status', # PROJECT PARAM
                   #'Designing Company', # YES
                   #'Project Manager MEP', # YES
                   #'Project Manager ID',# YES
                   #'Designer', # YES
                   #'Designer ID', # YES
                   #'Draftsperson', # YES
                   #'Sheet Name' # PROJECT PARAM
                   'Sheet Number', # PROJECT PARAM
                   #'Design', # YES
                   #'Design No.', # YES
                   #'Drawing Type',
                   #'Project Status',
                   'Sheet Issue Date',
                   'Revision',
                   ]
    for i,row in enumerate(excel_dict):
        count_total_rows += 1

        if row['SOURCE'] == 'RVT':

            count_total_rvt+=1
            
            # Use this to place a limit on the processing
            if count_total_rvt == 3 and 0:
                break
                        
            
            # Get sheet (exists)
            sheet_name = row['Sheet Name']

            try: this_sheet = sheets_by_name[row['Sheet Name']]
            except:
                logging.error("Sheet {} does not exist (not loaded in dictionary)".format(sheet_name))
                raise
            
            print("Processing Revit Sheet {}".format(sheet_name))
            
            changed_params = list()
            changed_values = list()
                        
            for k in row:
                if k in focus_params:
                    # 
                    flg_exists = util_ra.parameter_exists(this_sheet, k)
                    
                    if flg_exists:
                        value = util_ra.get_parameter_plain_string(this_sheet, k)
                    else:
                        value = "DNE"
                    
                    if row[k] == value:
                        flg_match = True
                    else:
                        flg_match = False
                    
                    if verbose and 1:
                        print("Row {} {:20} Param:{} Value:{:20} New Value:{:20} Match:{}".format(i,
                                                                                               row['Sheet Name'],
                                                                                               k,
                                                                                               value,
                                                                                               row[k],
                                                                                               flg_match
                                                                                               ))
                              
                if k in focus_params and not flg_match:
                    #time.sleep(1)
                    #print("PAUSE")
                    #time.sleep(1)
                    #print("PAUSE")
                    #time.sleep(1)
                    #print("PAUSE")
                                
                    # The parameter in excel does not match Revit => Update
                    changed_params.append(k)
                    changed_values.append(row[k])
                    count_total_updated += 1
                    
                    if flg_update:
                        try:
                            util_ra.change_parameter(rvt_doc, this_sheet, k, row[k],verbose=True, flg_error_prone=True)
                        except: 
                            print("*** Skip ***".format())
                            count_total_failed += 1
                            failed_sheets.append(sheet_name)
                            pass
                    
            print(" -- updated row {}, {} {}:  {}-> {} ".format(i, sheet_name, changed_params,value,changed_values))
                        
        else:
            if verbose:
                print("Skipped {}, row {}".format(row['SOURCE'],i))
                continue
                            
            #util_ra.table_parameters(this_sheet)
            
    logging.debug("Iterated {} document rows".format(count_total_rows))
    logging.debug("Checked {} sheets_by_name (RVT plan views only)".format(count_total_rvt))
    logging.debug("Updated {} sheets".format(count_total_updated-count_total_failed))
    logging.debug("Transaction failure on {} sheets".format(count_total_failed))
    for s in failed_sheets:
        print(" -- {}".format(s))

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
name_csv = r"\20161108 Document Register.csv"
path_csv = folder_csv + name_csv

#-Get data---
data_dict = util_ra.utility_general.get_data_csv(path_csv)

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
update_sheet_parameters(data_dict_RVT,sheets_by_name,flg_update=True,verbose=True)

logging.info("---DONE---".format())