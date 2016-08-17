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

#path_package = r"C:\Dropbox\00 CAD Standards\70 Revit Python\Packages\xlrd"
#sys.path.append(path_package)
#import xlrd

path_package = r"C:\EclipseGit\ExergyUtilities\ExergyUtilities"
sys.path.append(path_package)
import utility_excel_api_Interop as util_excel
import utility_revit_api as util_ra

#===============================================================================
# Definitions
#===============================================================================

def get_excel(path_excel_book):
    #-Get Excel data---
    with util_excel.ExtendedExcelBookAPI(path_excel_book) as xl:
        data_table = xl.get_table('REGISTER', 2, 200,1,42)

    header = data_table.pop(0)  
    table_dict = list()
    for row in data_table:
        if row[0]:
            table_dict.append(dict(zip(header,row)))
            
    logging.debug("Loaded {} sheet definitions".format(len(table_dict)))
    return table_dict


def check_sheets_exist(excel_dict, sheets):
    count_missing = 0
    count_total = 0
    for row in excel_dict:
        if row['SOURCE'] == 'RVT':
#             print("{} {}".format(row['NUMBER ON SHEET'],
#                                   row['NAME ON SHEET']))
#             
            try:
                sheets[row['NAME ON SHEET']]
                count_total += 1
            except:
                logging.error("Missing sheet:{} {}".format(row['NUMBER ON SHEET'],row['NAME ON SHEET']))
                count_missing += 1
        if count_missing: 
            raise Exception("{} missing sheets".format(count_missing))
        
    logging.debug("Checked {} sheets, all exist".format(count_total))

def check_sheets_views(excel_dict, sheets, views):
    count_missing = 0
    count_total = 0
    for i,row in enumerate(excel_dict):
        if row['SOURCE'] == 'RVT' and row['VIEW TYPE'] == 'PLAN':
#             try:
            sheet_name = row['NAME ON SHEET']
            sheet_num = row['NUMBER ON SHEET']
            main_view_name = row['VIEW NAME']
            legend = row['LEGEND']
            location = row['LOCATION LEGEND']
            paper_size = row['PAPER SIZE']
#             except:
#                 logging.error("Missing element:{}".format(row))
#                 
            logging.info("Processing row {}, {} {}".format(i,sheet_num,sheet_name))
            logging.info("\tMain view: {}".format(main_view_name))
            logging.info("\tLegend: {}".format(legend))
            logging.info("\tlocation view: {}".format(location))
            logging.info("\tpaper size: {}".format(paper_size))
            
            try: 
                sheets[row['NAME ON SHEET']]
            except:
                logging.error("Sheet {} does not exist".format(row['NAME ON SHEET']))
                raise
                
            try:
                views[row['VIEW NAME']]
            except:
                logging.error("View {} does not exist".format(row['VIEW NAME']))
                raise                
    raise           

    logging.debug("Done".format(count_total))


#===============================================================================
# Main
#===============================================================================
#-Logging info---
logging.info("Python version : {}".format(sys.version))
logging.info("uidoc : {}".format(rvt_uidoc))
logging.info("doc : {}".format(rvt_doc))
logging.info("app : {}".format(rvt_app))

#-Paths---
folder_excel_book = r"C:\CesCloud Revit\_03_IKEA_Working_Folder"
name_excel_book = r"\20160629 Document Register.xlsx"
path_excel_book = folder_excel_book + name_excel_book

#-Get Excel data---
excel_dict = get_excel(path_excel_book)

#-Get all floorplans, sheets, titleblocks, legends---
util_ra.get_all_views(rvt_doc)
title_blocks = util_ra.get_title_blocks(rvt_doc)
sheets = util_ra.get_sheet_dict(rvt_doc)
floorplans = util_ra.get_views_by_type(rvt_doc, 'FloorPlan')
legends = util_ra.get_views_by_type(rvt_doc,'Legend')
templates = util_ra.get_view_templates(rvt_doc)

#-Check---
check_sheets_exist(excel_dict,sheets)
check_sheets_views(excel_dict,sheets,floorplans)
raise

# Create sheet---
sheet_name = 'TEST SHEET NAME'
sheet_number = 'Test number2'
this_title_block = title_blocks['A0 IKEA Title 1_250']
this_sheet = util_ra.create_sheet(rvt_doc, this_title_block, sheet_name, sheet_number)
#this_sheet = sheets[sheet_name]

# Add view---
view_name = '00 - GROUND FLOOR LEVEL Copy 1'
this_view = floorplans[view_name]
util_ra.add_view_sheet(rvt_doc, this_sheet, this_view, OVERVIEWPLAN_CENTER)

logging.info("---DONE---".format())


#xl_sheet = xl_workbook.sheet_by_index(0)
#logging.info("excel sheet : {}".format(xl_sheet))