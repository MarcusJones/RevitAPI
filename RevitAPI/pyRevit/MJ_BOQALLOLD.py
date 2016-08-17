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

#===============================================================================
# Imports other
#===============================================================================
import sys
import csv
import time

path_package = r"C:\EclipseGit\ExergyUtilities\ExergyUtilities"
sys.path.append(path_package)
import utility_revit_api as util_ra


#===============================================================================
# Definitions
#===============================================================================

#===============================================================================
# Main
#===============================================================================
#-Logging info---
logging.info("Python version : {}".format(sys.version))
logging.info("uidoc : {}".format(rvt_uidoc))
logging.info("doc : {}".format(rvt_doc))
logging.info("app : {}".format(rvt_app))
 
"""This is a dict of only INSTANCES in project"""
#el_dict_instances = util_ra.get_sort_all_FamilyInstance(rvt_doc)
start = time.time()
table = list()
for i,elem in enumerate(util_ra.get_all_FamilyInstance(rvt_doc)):
    print(elem)
    print(elem.Id)
    #print(elem.Id.Value)
    print(elem.Category)
    print(elem.Category.ToString())
    print(elem.Category.Name)
    print(elem.Category.Parent)
    print(elem.Category.Id)
    print("****")
    util_ra.get_elem_BuiltInCategory(elem)
    raise
    
    this_row = dict()
    
    # Classificaitons
    
    this_row["Category"] = elem.Category.Name
    
    ws_table = rvt_doc.GetWorksetTable()
    this_ws = ws_table.GetWorkset(elem.WorksetId)
    this_row["Workset"] = this_ws.Name
    
    this_row["Family"] = elem.Symbol.FamilyName
    
    this_row["Type"] = elem.Name
    
    this_row["Size"] = util_ra.get_parameter_value(elem, "Size", flg_DNE=True)

    this_row["System"] = util_ra.get_parameter_value(elem, "System Abbreviation", flg_DNE=True)

    # Report 
    
    this_row["Description"] = util_ra.get_parameter_value(elem, "Description", flg_DNE=True)

    this_row["ifcDescription"] = util_ra.get_parameter_value(elem, "ifcDescription", flg_DNE=True)
    
    this_row["Length"] = util_ra.get_parameter_value_float(elem, "Length", flg_DNE=True)
    
    this_row["Area"] = util_ra.get_parameter_value_float(elem, "Area", flg_DNE=True)
    
    this_row["Id"] = util_ra.get_parameter_value(elem, "Id", flg_DNE=True)
    
    #this_row["IKEA Item Code"] = util_ra.get_parameter_value(elem, "IKEA Item Code")
    
    #this_row["IKEA Cost Group"] = util_ra.get_parameter_value(elem, "IKEA Cost Group")
    
    # Enforce strings
    for k in this_row:
        try:
            this_row[k] = this_row[k].encode('ascii','ignore')
        except:
            continue
            #print(this_row)
            #raise
        
    table.append(this_row)
    #print(this_row)
    #if i == 10000:
    #    break  

#print(util_ra.format_as_table(table, table[0].keys(),table[0].keys()))


end = time.time()
        
logging.info("Built table of {} elements over {} seconds".format(len(table),end - start))
    


out_path = r"C:\CesCloud Revit\_03_IKEA_Working_Folder\BOQ Development\BOQ ALL.csv"
with open(out_path,'wb') as csv_file:
    writer = csv.DictWriter(csv_file, delimiter=';', fieldnames=table[0].keys())
    writer.writeheader()
    writer.writerows(table)
    
end = time.time()
        
logging.info("Wrote table to {} over {} seconds".format(out_path,end - start))
    

#for cat_string in util_ra.REVIT_CATEGORIES:
#    print("{:30} : {:5} elements".format(cat_string, len(el_dict_instances[cat_string])))
# 
# # Build a set of all FamilySymbol names which do have placed instances
# placed_type_names = set()
# placed_type_ids = set()
# #place
# for cat_string in util_ra.REVIT_CATEGORIES:
#     for instance in el_dict_instances[cat_string]:
#         placed_type_ids.add(instance.GetTypeId())
# 
# table = list()
# for i,type_id in enumerate(placed_type_ids):
#     this_row = dict()
#     this_type = rvt_doc.GetElement(type_id)
#     elems = util_ra.select_instances_by_type_id(rvt_doc, type_id)
#     this_el = elems[0]
#     assert util_ra.parameter_exists(this_type, "ifcDescription")
#     
#     for ws in workset:
#     
#     this_row["QTY"] = len(elems)
#     this_row["Family"] = this_type.FamilyName
#     this_row["Type"] = elems[0].Name
#     this_row["ifcDescription"] = util_ra.get_parameter_value(this_type, "ifcDescription")
#     this_row["Category"] = this_type.Category.Name
#     
#     
#     #util_ra.all_params(this_type)
#     table.append(this_row)
#     #print(this_row)
# 
# 
# out_path = r"C:\CesCloud Revit\_03_IKEA_Working_Folder\BOQ Development\BOQ.csv"
# with open(out_path,'wb') as csv_file:
#     writer = csv.DictWriter(csv_file, delimiter=';', fieldnames=table[0].keys())
#     writer.writeheader()
#     writer.writerows(table)
#     
#     #raise
#     #
#     #row_string = ";".join([str(len(elems)), ";", elem.FamilyName, this_el.Name])
#     #table.append(row_string)




# for row in table:
#     print(table)
#     #print()


#print(len(placed_type_ids))
#print()
#raise

#for i,type in enumerate(placed_types):
#    print(i,type)
#    #print(type.Name, type)

#for type_name in placed_types:
#    print("{}".format(sym))





#print(util_ra.family_data_dict(rvt_doc,instance))        

#print(placed_types)

        
if 0:
    for fam in el_dict_instances['Communication Devices']:
        assert type(fam) == rvt_db.FamilyInstance, "Only placed instances"
        print(util_ra.family_data_dict(rvt_doc,fam))
    
    print("**********")
    print(fam)
    print(util_ra.all_params(fam))    
    
    #print(el)
    #print(type(el))
    #print_family
    #util_ra.print_family(el)
    
logging.info("---DONE---".format())
