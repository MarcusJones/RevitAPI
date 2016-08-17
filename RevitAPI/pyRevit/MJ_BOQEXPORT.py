from __future__ import print_function

#===============================================================================
# pyRevit
#===============================================================================
__doc__ = """

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
 
start = time.time()
table = list()

all_instances = util_ra.get_BOQ_elements(rvt_doc)

for i,elem in enumerate(all_instances):

    if i % 100 == 0:
        print("{} of {}".format(i,len(all_instances)))
    this_row = dict()
    
    # Classificaitons
    try:
        this_row["Category"] = elem.Category.Name
    except:
        print(elem)
        print(elem.Name)
        print(elem.Category)
        #continue
        raise
    
    try:
        ws_table = rvt_doc.GetWorksetTable()
        this_ws = ws_table.GetWorkset(elem.WorksetId)
        this_row["Workset"] = this_ws.Name
    except:
        #print("Workset problem",elem)
        this_row["Workset"] = "UNDEFINED"
        pass
    
    try:
        this_row["Family"] = elem.Symbol.FamilyName
    except:
        #print("Family problem",elem)
        this_row["Family"] = "UNDEFINED"
        pass
    
    try:
        this_row["Type"] = elem.Name
    except:
        #print("Type problem",elem)
        this_row["Type"] = "UNDEFINED"            
        pass
    
    this_row["Size"] = util_ra.get_parameter_value(elem, "Size", flg_DNE=True)

    this_row["System"] = util_ra.get_parameter_value(elem, "System Abbreviation", flg_DNE=True)

    # Report 
    
    this_row["Description"] = util_ra.get_parameter_value(elem, "Description", flg_DNE=True)
    
    this_row["ifcDescription"] = util_ra.get_parameter_value(elem, "ifcDescription", flg_DNE=True)
    
    this_row["Length"] = util_ra.get_parameter_value_str(elem, "Length", flg_DNE=True)
    
    this_row["Area"] = util_ra.get_parameter_value_str(elem, "Area", flg_DNE=True)
    
    this_row["Id"] = elem.Id
    
    this_row["IKEA Item Code"] = util_ra.get_parameter_value(elem, "IKEA Item Code", flg_DNE=True)
    
    this_row["IKEA Cost Group"] = util_ra.get_parameter_value(elem, "IKEA Cost Group", flg_DNE=True)
    
    # Enforce strings
    for k in this_row:
        try:
            this_row[k] = this_row[k].encode('ascii','ignore')
        except:
            #logging.error("Problem encoding {} = {} {} of elem {}".format(k,this_row[k], type(this_row[k]), elem.Name))
            #print(sys.exc_info()[0])
            continue

    table.append(this_row)

#print(util_ra.format_as_table(table, table[0].keys(),table[0].keys()))

end = time.time()
        
logging.info("Built table of {} elements over {} seconds".format(len(table),end - start))
    
out_path = r"C:\CesCloud Revit\_03_IKEA_Working_Folder\BOQ Development\BOQ ALL3.csv"
with open(out_path,'wb') as csv_file:
    writer = csv.DictWriter(csv_file, delimiter=';', fieldnames=table[0].keys())
    writer.writeheader()
    writer.writerows(table)
    
end = time.time()
        
logging.info("Wrote table to {} over {} seconds".format(out_path,end - start))
    
logging.info("---DONE---".format())


"""
categories = rvt_doc.Settings.Categories;
#print("{:40} | {:30} | {:30}".format("cat.Name", "cat.Id", "cat.Parent"))
cat_ids = list()    
for cat in categories:
    if cat.Name in util_ra.REVIT_CATEGORIES:
        cat_ids.append(cat.Id)


#util_ra.print_dir(rvt_db.BuiltInCategory)
print(rvt_db.BuiltInCategory)
raise
target_bics = list()
for bic in rvt_db.BuiltInCategory:
    for c_id in cat_ids:
        if c_id == bic.Id:
            target_bics.append(bic)

for bi in target_bics:
    print(bi)

#this_cat = rvt_db.BuiltInCategory(cat_ids[0])
#print(this_cat)
        #print(cat.Name)
    #else:
    #    print("NO:",cat.Name)
    #print("{:40} | {:30} | {:30}".format(cat.Name, cat.Id, cat.Parent))

#print(rvt_db.BuiltInCategory("Duct"))
raise
all_instance_IDs = util_ra.get_all_elements_IDs(rvt_doc)
#all_mep_IDs = util_ra.get_all_MEP_element_IDs(rvt_doc)

#print(all_instances[0])
#raise
#all_instances_ids = [el.Id for el in all_instances]
#all_mep_ids = [el.Id for el in all_mep]
# 
# for mep_el_id in all_mep_IDs:
#     if mep_el_id not in all_instance_IDs:
#         all_instance_IDs.append(mep_el_id)
#     else:
#         continue            


all_instances_merged = [util_ra.get_element_by_id(rvt_doc,el_id) 
                    for el_id in all_instance_IDs]


# 
# logging.info("Merged {} MEP elements, bringing total from {} to {}".format(len(all_mep_IDs),
#                                                                            len(all_instance_IDs),
#                                                                            len(all_instances_merged)))

#raise

# for el in all_instances_merged:
#     print(el)
"""

"""This is a dict of only INSTANCES in project"""
#el_dict_instances = util_ra.get_sort_all_FamilyInstance(rvt_doc)



"""

#     print(i)
#     print(elem)
#     print(elem.Id)
#     #print(elem.Id.Value)
#     print(elem.Category)
#     print(elem.Category.ToString())
#     print(elem.Category.Name)
#     print(elem.Category.Parent)
#     print(elem.Category.Id)
#     print("****")
#     util_ra.get_elem_BuiltInCategory(rvt_doc,elem)
#     raise
    #print("{} of {}".format(i,total))    
"""