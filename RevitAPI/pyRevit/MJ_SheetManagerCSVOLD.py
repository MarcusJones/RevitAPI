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
def get_data_csv(path_csv, this_delimiter=';'):
    table_dict = list()
    with open(path_csv) as csvfile:
        
        # First, open the file to get the header, skip one line
        reader = csv.reader(csvfile,delimiter=this_delimiter)
        skip_row = next(reader)
        headers = next(reader)
        #print(headers)
        #print(type(headers))
        #raise
        # Use the header, re-read, and skip 2 lines
        reader = csv.DictReader(csvfile,fieldnames=headers,delimiter=';')
        skip_row = next(reader)
        skip_row = next(reader)
        
        for row in reader:
            #print("ROW START")
            op_A = False
            for k in row:
                found = op_A or bool(row[k]) 
                if found: 
                    break
            if not found:
                break
            #print(row)
            table_dict.append(row)
                #if bool(row[k]):
                    
                #    break
                #print(row[k], bool(row[k]))
            
            
            #print(row)
    #print(reader[3])
    logging.debug("Loaded {} sheet definitions with {} columns".format(len(table_dict), len(table_dict[0])))
    
    return table_dict

def update_sheet_parameters(excel_dict, sheets_by_name, flg_update=False, verbose=False):
    logging.debug(util_ra.get_self())
    count_total_rows = 0
    count_total_rvt = 0
    focus_params =[
                   #'Project Number', # PROJECT PARAM
                   #'Project Status', # PP
                   'Designing Company', # YES
                   'Project Manager MEP', # YES
                   'Project Manager ID',# YES
                   'Designer', # YES
                   'Designer ID', # YES
                   'Draftsperson', # YES
                   #'Sheet Name' # PP
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

def create_sheets(data_dict, sheets_by_name, title_block):
    """
    Create a new blank sheet with NAME and NUMBER on TITLEBLOCK
    """
    logging.debug(util_ra.get_self())
    
    #assert "TEMPLATE" in sheets_by_name, "Template does not exist"
    
    #template_sheet = sheets_by_name["TEMPLATE"]
    cnt_new = 0
    #print(sheets_by_name)
    for i,row in enumerate(data_dict):
        assert row['SOURCE'] == 'RVT', "Only works with RVT drawings, not [{}]".format(row['SOURCE'])
        if row['Sheet Name'] not in sheets_by_name:
            #print("CREATE:",row['Sheet Name'])
            #assert row["View Name"] in views, "View {} does not exist"
            util_ra.create_sheet(rvt_doc, title_block, row['Sheet Number'], row['Sheet Name'])
            cnt_new+=1
    logging.debug("Created {} sheets".format(cnt_new))


def rename_sheets(data_dict, sheets_by_name):
    raise "OBSELETE SEE CREATE SHEETS"    
    logging.debug(util_ra.get_self())
    
    for i,row in enumerate(data_dict):
        assert row['SOURCE'] == 'RVT', "Only works with RVT drawings, not [{}]".format(row['SOURCE'])
        if row['OLD NAME']:
            assert row['OLD NAME'] in sheets_by_name, "{} sheet not found in project".format(row['OLD NAME'])
        else:
            continue
        
        this_sheet = sheets_by_name[row['OLD NAME']]
        
        # Change OLD NAME to Sheet Name column
        util_ra.change_parameter(rvt_doc, 
                                 this_sheet, 
                                 'Sheet Name', 
                                 row['Sheet Name'])
        
        # Change OLD NUMBER to Sheet Number column
        util_ra.change_parameter(rvt_doc, 
                                 this_sheet, 
                                 'Sheet Number', 
                                 row['Sheet Number'])        
        
        logging.debug("Updated {}".format(row['OLD NAME']))

def check_sheets(excel_dict, sheets_by_name, views):
    logging.debug(util_ra.get_self())
    
    count_total_rows = 0
    count_total_rvt = 0
    count_missing_view = 0
    count_missing_sheet = 0
    count_missing_view = 0
    for i,row in enumerate(excel_dict):
        count_total_rows += 1
        assert row['SOURCE'] == 'RVT'
        if row['VIEW TYPE'] == 'PLAN':
            count_total_rvt += 1

            logging.info("CHECK: {} {} {} View: {} Legend: {} Loc: {}".format(row['Sheet Name'],
                                     row['Sheet Number'],
                                     row['PAPER SIZE'],
                                     row['View Name'],
                                     row['LEGEND'],
                                     row['LOCATION LEGEND']
                                     ))
           
            # Test sheet exists
            if row['Sheet Name'] not in sheets_by_name:
                count_missing_sheet+=1
                logging.error("Sheet {} does not exist".format(row['Sheet Name']))
                
            
            # Test view exists
            if row['View Name'] not in views:
                count_missing_view
                logging.error("View {} does not exist".format(row['View Name']))
                                
            # Test view in register matches view in Revit over sheet placed views
            if (row['Sheet Name'] not in sheets_by_name and 
                row['View Name'] not in views):
                placed_views = sheets_by_name[row['Sheet Name']].GetAllPlacedViews() 
                view_match = False
                for view_id in placed_views:
                    view = rvt_doc.GetElement(view_id)
                    if view.Name == row['View Name']:
                        view_match = True
                
                #assert view_match
                if not view_match:
                    count_missing_view+=1
                    logging.error("View {} not placed on {}".format(row['Sheet Name'], row['View Name']))

    logging.debug("Iterated {} document rows".format(count_total_rows))
    logging.debug("Checked {} sheets_by_name (RVT plan views only)".format(count_total_rvt))
    logging.debug("Missing sheets: {} ".format(count_missing_sheet))
    logging.debug("Missing views: {}".format(count_missing_view))
    logging.debug("{} placed sheets had a missing or unmatched view".format(count_missing_view))    
    

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
data_dict = get_data_csv(path_csv)

# data_dict_RVT = list()
# for row in data_dict:
#     for i,row in enumerate(excel_dict):
#         count_total_rows += 1
#         if row['SOURCE'] == 'RVT' and row['VIEW TYPE'] == 'PLAN':
#             data_dict_RVT.append()
data_dict_RVT = [row for row in data_dict if row['SOURCE'] == 'RVT']

# START WITH ONE ROW
#data_dict = data_dict[0:8]

#-Get all floorplans, sheets_by_name, titleblocks, legends---
util_ra.get_all_views(rvt_doc)
title_blocks = util_ra.get_title_blocks(rvt_doc)
sheets_by_name = util_ra.get_sheet_dict_by_names(rvt_doc)
floorplans = util_ra.get_views_by_type(rvt_doc, 'FloorPlan')
all_views = util_ra.get_all_views(rvt_doc)
legends = util_ra.get_views_by_type(rvt_doc,'Legend')
templates = util_ra.get_view_templates(rvt_doc)

#-Check---
print("*****************************")
#check_sheets(data_dict_RVT,sheets_by_name,floorplans)
print("*****************************")
#update_sheet_parameters(data_dict_RVT,sheets_by_name,flg_update=False)
print("*****************************")

#-Create sheet---
create_sheets(data_dict_RVT,sheets_by_name,title_blocks['A0 IKEA Title 1_250'])







#update_sheets_views(data_dict_RVT,sheets_by_name,floorplans,flg_update=False)


#sheet_name = 'TEST SHEET NAME'
#sheet_number = 'Test number2'
#this_title_block = title_blocks['A0 IKEA Title 1_250']
#this_sheet = util_ra.create_sheet(rvt_doc, this_title_block, sheet_name, sheet_number)
#this_sheet = sheets_by_name[sheet_name]

#-Add view---
#view_name = '00 - GROUND FLOOR LEVEL Copy 1'
#this_view = floorplans[view_name]
#util_ra.add_view_sheet(rvt_doc, this_sheet, this_view, OVERVIEWPLAN_CENTER)

logging.info("---DONE---".format())


#xl_sheet = xl_workbook.sheet_by_index(0)
#logging.info("excel sheet : {}".format(xl_sheet))