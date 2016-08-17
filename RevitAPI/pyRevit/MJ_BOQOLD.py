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
 

#util_ra.get_all_categories(rvt_doc)

#print(rvt_doc.Settings)
#util_ra.print_dir(rvt_doc.Settings)

#print(rvt_doc.Settings.Categories.Item('Air Terminals'))

"""This is a dict of all elements, including line style, FamilySymbol, etc. """
#el_dict_all = util_ra.get_sort_all_elements(rvt_doc)
#for cat_string in util_ra.REVIT_CATEGORIES:
#    print("{:30} : {:5} elements".format(cat_string, len(el_dict_all[cat_string])))

"""This is a dict of only INSTANCES in project"""
el_dict_instances = util_ra.get_sort_all_FamilyInstance(rvt_doc)

for cat_string in util_ra.REVIT_CATEGORIES:
    print("{:30} : {:5} elements".format(cat_string, len(el_dict_instances[cat_string])))

# Build a set of all FamilySymbol names which do have placed instances
placed_type_names = set()
placed_type_ids = set()
#place
for cat_string in util_ra.REVIT_CATEGORIES:
    for instance in el_dict_instances[cat_string]:
        placed_type_ids.add(instance.GetTypeId())

table = list()
for i,type_id in enumerate(placed_type_ids):
    this_row = dict()
    this_type = rvt_doc.GetElement(type_id)
    elems = util_ra.select_instances_by_type_id(rvt_doc, type_id)
    this_el = elems[0]
    assert util_ra.parameter_exists(this_type, "ifcDescription")
    
    for ws in workset:
    
    this_row["QTY"] = len(elems)
    this_row["Family"] = this_type.FamilyName
    this_row["Type"] = elems[0].Name
    this_row["ifcDescription"] = util_ra.get_parameter_value(this_type, "ifcDescription")
    this_row["Category"] = this_type.Category.Name
    
    
    #util_ra.all_params(this_type)
    table.append(this_row)
    #print(this_row)


out_path = r"C:\CesCloud Revit\_03_IKEA_Working_Folder\BOQ Development\BOQ.csv"
with open(out_path,'wb') as csv_file:
    writer = csv.DictWriter(csv_file, delimiter=';', fieldnames=table[0].keys())
    writer.writeheader()
    writer.writerows(table)
    
    #raise
    #
    #row_string = ";".join([str(len(elems)), ";", elem.FamilyName, this_el.Name])
    #table.append(row_string)




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
