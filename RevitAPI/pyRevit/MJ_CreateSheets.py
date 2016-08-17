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

def create_sheets(data_dict, sheets_by_name, title_block):
    """
    Create a new blank sheet with NAME and NUMBER on TITLEBLOCK
    """
    logging.debug(util_ra.get_self())
    
    #assert "TEMPLATE" in sheets_by_name, "Template does not exist"
    
    sheet_numbers = [sheets_by_name[sheet_name].SheetNumber for sheet_name in sheets_by_name]
    
    #template_sheet = sheets_by_name["TEMPLATE"]
    cnt_new = 0
    #print(sheets_by_name)
    for i,row in enumerate(data_dict):
        assert row['SOURCE'] == 'RVT', "Only works with RVT drawings, not [{}]".format(row['SOURCE'])
        if row['Sheet Name'] not in sheets_by_name:
            if row['Sheet Number'] in sheet_numbers:
                logging.error("Sheet number {} already exists".format(row['Sheet Number']))
                raise    
                    
            #print("CREATE:",row['Sheet Name'])
            #assert row["View Name"] in views, "View {} does not exist"
            util_ra.create_sheet(rvt_doc, title_block, row['Sheet Number'], row['Sheet Name'])
            cnt_new+=1
    logging.debug("Created {} sheets".format(cnt_new))


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
title_blocks = util_ra.get_title_blocks(rvt_doc)
sheets_by_name = util_ra.get_sheet_dict_by_names(rvt_doc)

#-Create sheet---
create_sheets(data_dict_RVT,sheets_by_name,title_blocks['A0 IKEA Title 1_250'])

logging.info("---DONE---".format())
