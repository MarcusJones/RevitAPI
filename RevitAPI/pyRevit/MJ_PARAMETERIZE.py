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

def get_keyed_data(this_table):
    # Refactor the table to a lookup table with compound key
    # [(family, type)] = ...
    keyed_table = dict()
    for row in this_table:
        
        this_key = (row['Workset'],
            row['Category'],
            row['Family'],
            row['Type'],
            row['System'],
            row['Size'])
        
        this_val = {'Item Code':row['IKEA Item Code NEW'],
                    'ifcDescription':row['ifcDescription NEW'],
                    'Cost Group':row['IKEA Cost Group NEW']
                    }
        keyed_table[this_key] = this_val
        #print(this_key)
    logging.debug("Grouped into {} rows".format(len(keyed_table)))
    
    return keyed_table


def get_match(family, type, all_instances):
    raise


def parameterize_ver2(names_dict):
    logger = logging.getLogger()
    all_instances = util_ra.get_BOQ_elements(rvt_doc)
    ws_table = rvt_doc.GetWorksetTable()
    count_fnd = 0
    count_total = 0
    count_missed_item_code = 0
    count_missed_ifcDescription = 0    
    unmatched_elems = list()
    for i,elem in enumerate(all_instances):
        count_total += 1
        # Check if element is valid
        MEP_elems = [
            rvt_db.Electrical.CableTray,                         
            rvt_db.Electrical.Wire,                        
            rvt_db.Electrical.Conduit,                           
            rvt_db.Mechanical.Duct,                         
            rvt_db.Plumbing.Pipe,                      
            ]
        if type(elem) == rvt_db.FamilyInstance:
            pass
            #print("Instance")
        elif type(elem) in MEP_elems:
            pass
            #print("MEP")
        else:
            print(i,elem)
            raise        
            
        this_elem = dict()
        
        # Workset ***
        this_ws = ws_table.GetWorkset(elem.WorksetId)
        this_elem["Workset"] = this_ws.Name
        
        # Category ***
        this_elem["Category"] = elem.Category.Name
        
        # Family ***
        try:
            this_elem["Family"] = elem.Symbol.FamilyName
            this_elem["Family"] = this_elem["Family"].encode('ascii','ignore')
        except:
            #print("Family problem",elem)
            this_elem["Family"] = "UNDEFINED"
            pass
                
        # Type ***
        try:
            this_elem["Type"] = elem.Name
            this_elem["Type"] = this_elem["Type"].encode('ascii','ignore')
        except:
            #print("Type problem",elem)
            this_elem["Type"] = "UNDEFINED"            
            pass

        # System ***
        this_elem["System"] = util_ra.get_parameter_value(elem, "System Abbreviation", flg_DNE=True)
        this_elem["System"] = this_elem["System"].encode('ascii','ignore')
        
        # Size ***
        this_elem["Size"] = util_ra.get_parameter_value(elem, "Size", flg_DNE=True)
        this_elem["Size"] = this_elem["Size"].encode('ascii','ignore')

        #this_elem["Size"] = util_ra.get_parameter_value(elem, "Size", flg_DNE=True)

        search_key = (this_elem["Workset"],
                      this_elem["Category"],
                      this_elem["Family"],
                      this_elem["Type"],
                      this_elem["System"],
                      this_elem["Size"],
                      )

        if search_key in names_dict:
            count_fnd += 1

            
            
            logger.setLevel(logging.INFO)
            try:
                util_ra.change_parameter(rvt_doc, elem, 'IKEA Cost Group', names_dict[search_key]['Cost Group'],verbose = False)
            except:
                raise
                count_missed_item_code += 1

            try:
                util_ra.change_parameter(rvt_doc, elem, 'ifcDescription', names_dict[search_key]['ifcDescription'],verbose = False)
            except:
                count_missed_item_code
                count_missed_ifcDescription += 1
                pass
            logger.setLevel(logging.DEBUG)
            
        else: 
            #print(search_key)
            unmatched_elems.append(this_elem)
            pass
        
        start_range = 0
        end_range = 10000
        
        if i > start_range and i < end_range:
            break
        
        if i % 1 == 0:
            print(i,search_key)
            #print(names_dict[search_key])
            
    print("Matched {} elements out of {}".format(count_fnd, count_total))
    print("Out of {} matches, Cost Group failed {} times".format(count_fnd,count_missed_item_code))
    print("Out of {} matches, ifcDescription failed {} times".format(count_fnd,count_missed_ifcDescription))
    
    #for unmatched in unmatched_elems:
    #    print(unmatched)
    
MEP_ELEMENTS_TYPES = [
    rvt_db.Electrical.CableTray,                         
    rvt_db.Electrical.Wire,                        
    rvt_db.Electrical.Conduit,                           
    rvt_db.Mechanical.Duct,                         
    rvt_db.Plumbing.Pipe,                      
    ]
    
def element_valid(elem):
            
    if type(elem) == rvt_db.FamilyInstance:
        #print("Instance")
        return True
    
    elif type(elem) in MEP_ELEMENTS_TYPES:
        pass
        return True
    else:
        print(elem)
        return False



def get_elem_info(elem):
    this_elem_dict = dict()
    
    # Workset ***
    this_ws = WORKSET_TABLE.GetWorkset(elem.WorksetId)
    this_elem_dict["Workset"] = this_ws.Name
    
    # Category ***
    this_elem_dict["Category"] = elem.Category.Name
    
    # Family ***
    try:
        this_elem_dict["Family"] = elem.Symbol.FamilyName
        this_elem_dict["Family"] = this_elem_dict["Family"].encode('ascii','ignore')
    except:
        #print("Family problem",elem)
        this_elem_dict["Family"] = "UNDEFINED"
        pass
            
    # Type ***
    try:
        this_elem_dict["Type"] = elem.Name
        this_elem_dict["Type"] = this_elem_dict["Type"].encode('ascii','ignore')
    except:
        #print("Type problem",elem)
        this_elem_dict["Type"] = "UNDEFINED"            
        pass

    # System ***
    this_elem_dict["System"] = util_ra.get_parameter_value(elem, "System Abbreviation", flg_DNE=True)
    this_elem_dict["System"] = this_elem_dict["System"].encode('ascii','ignore')
    
    # Size ***
    this_elem_dict["Size"] = util_ra.get_parameter_value(elem, "Size", flg_DNE=True)
    this_elem_dict["Size"] = this_elem_dict["Size"].encode('ascii','ignore')

    #this_elem_dict["Size"] = util_ra.get_parameter_value(elem, "Size", flg_DNE=True)
    
    return this_elem_dict
     
def parameterize_ver3(names_dict):
    all_instances = util_ra.get_BOQ_elements(rvt_doc)
    count_fnd = 0
    count_missed = 0
    count_total = 0
    count_skip = 0
    block_size = 1000
    
    el_blocks = list()
    
    this_el_list = list()
    
    for i,elem in enumerate(all_instances):
        this_el_list.append(elem)
        if i % block_size == 0 and i != 0:
            # Add this block
            el_blocks.append(this_el_list)
            # Reset the block
            this_el_list = list()
            
    for i,block in enumerate(el_blocks):
        logging.debug("Block {} of size {}".format(i, len(block)))
        this_elem_list = list()
        new_val_list_COSTGROUP = list()
        new_val_list_IFCDESC = list()
        
        # Build the block
        for i,elem in enumerate(block):
            count_total += 1
                
            # Check if element is valid
            assert element_valid, "Not valid element"
            
            # Get the signature of this element
            try:
                this_elem_dict = get_elem_info(elem)
            except:
                print(elem)
                raise
            
            try:
                search_key = (this_elem_dict["Workset"],
                              this_elem_dict["Category"],
                              this_elem_dict["Family"],
                              this_elem_dict["Type"],
                              this_elem_dict["System"],
                              this_elem_dict["Size"],
                              )
            except:
                print("Search key problem")
                raise
            
            # Add this element to the lists
            if search_key in names_dict:
                
                flg_change_CG = util_ra.get_parameter_value(elem,'IKEA Cost Group') != names_dict[search_key]['Cost Group']
                flg_change_IFCD = util_ra.get_parameter_value(elem,'ifcDescription') != names_dict[search_key]['ifcDescription']
                
                if flg_change_CG or flg_change_IFCD:
                    count_fnd += 1
                    this_elem_list.append(elem)
                    new_val_list_COSTGROUP.append(names_dict[search_key]['Cost Group'])
                    new_val_list_IFCDESC.append(names_dict[search_key]['ifcDescription'])
                else:
                    count_skip += 1
            else: 
                count_missed += 1

        # Done building block, now commit
        logging.debug("Matched {} elements out of {}, {} Skipped no change, {} elements missing signatures in excel".format(count_fnd,count_total,count_skip,                                                                                                                                              count_missed))
        
        # Commit 
        util_ra.change_parameter_multiple(rvt_doc, this_elem_list, 'IKEA Cost Group', new_val_list_COSTGROUP)
        util_ra.change_parameter_multiple(rvt_doc, this_elem_list, 'ifcDescription', new_val_list_IFCDESC)

    
    #print("Out of {} matches, Cost Group failed {} times".format(count_fnd,count_missed_item_code))
    #print("Out of {} matches, ifcDescription failed {} times".format(count_fnd,count_missed_ifcDescription))
    
    #for unmatched in unmatched_elems:
    #    print(unmatched)
    
    logging.debug("{} blocks process".format(i))
    
def parameterize(names_dict):
    for i,row in enumerate(names_dict):
        #if i == 5:
        #    break
        
        logging.debug("{} PROCESSING {} = {}".format(i,row['Family'],row['Type']))
        target_family = row['Family']
        target_type = row['Type']
        
        #param_dict = names_dict[row]
        
        elems = rvt_db.FilteredElementCollector(rvt_doc).OfClass(rvt_db.FamilySymbol).ToElements()
        raise
        for i,el in enumerate(elems):
            this_category = None
            this_family = el.FamilyName
            this_type = util_ra.get_parameter_value(el,'Type Name')
            
            change_list = list()
            
            if target_family == this_family and target_type == this_type:
                
                # ITEM CODE ###########
                if util_ra.get_parameter_value(el,'IKEA Item Code') != row['IKEA Item Code NEW']:
                    
                    try:
                        util_ra.change_parameter(rvt_doc, el, 'IKEA Item Code', row['IKEA Item Code NEW'],verbose = False)
                        change_list.append("Item Code")
                    except:
                        print("Couldn't change {}".format('IKEA Item Code'))
                        change_list.append("ERROR in Item Code")
                        raise
                else:
                    change_list.append("Skipped Item Code")
                    
                # ifcDescription ###########
                if util_ra.get_parameter_value(el,'ifcDescription') != row['ifcDescription NEW']:
                    try:
                        util_ra.change_parameter(rvt_doc, el, 'ifcDescription', row['ifcDescription NEW'],verbose = False)
                        change_list.append("ifcDescription")
                    except:
                        print("Couldn't change {}".format('ifcDescription'))
                        change_list.append("ERROR in ifcDescription")
                        raise
                else:
                    change_list.append("Skipped ifcDescription")
                    
                # COST GROUP ###########                    
                if util_ra.get_parameter_value(el,'IKEA Cost Group') != row['IKEA Cost Group NEW']:
                    try:  
                        util_ra.change_parameter(rvt_doc, el, 'IKEA Cost Group', row['IKEA Cost Group NEW'],verbose = False)
                        change_list.append("IKEA Cost Group")
                        
                    except:
                        print("Couldn't change {}".format('IKEA Cost Group'))
                        change_list.append("ERROR in IKEA Cost Group")
                        raise
                else:
                    change_list.append("Skipped IKEA Cost Group")
            
            logging.debug("Changes: {}".format(";".join(change_list)))
                          
#-Paths---
folder_csv = r"C:\CesCloud Revit\_03_IKEA_Working_Folder\BOQ Development"
name_csv = r"\BOQ COMBINED1.csv"
path_csv = folder_csv + name_csv

#-Get keyed data---
#keyed_data = get_keyed_data(path_csv)
names_dict = util_ra.get_data_csv(path_csv)
keyed_data = get_keyed_data(names_dict)
#keyed_data = keyed_data[1:3] 

#print(keyed_data[10])
#-Parameterize elements---
#print(len(keyed_data))
#parameterize(names_dict[0:10])

WORKSET_TABLE = rvt_doc.GetWorksetTable()
parameterize_ver3(keyed_data)
#all_instances = util_ra.get_BOQ_elements(rvt_doc)

#parameterize(names_dict, all_instances)

logging.info("---DONE---".format())


if 0:    
    out_path = r"C:\CesCloud Revit\_03_IKEA_Working_Folder\BOQ Development\BOQ ALL2.csv"
    with open(out_path,'wb') as csv_file:
        writer = csv.DictWriter(csv_file, delimiter=';', fieldnames=table[0].keys())
        writer.writeheader()
        writer.writerows(table)
        
    end = time.time()
            
    logging.info("Wrote table to {} over {} seconds".format(out_path,end - start))
        

