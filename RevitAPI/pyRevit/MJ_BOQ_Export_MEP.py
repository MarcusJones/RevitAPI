from __future__ import print_function

#===============================================================================
# pyRevit
#===============================================================================
__doc__ = """
Get all INSTANCES of interest (MEP object)
Write a table to CSV of PRIMARY KEY and REPORTING data for all 

05.10.2016 - Update 
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

path_package = r"C:\EclipseGit\ExergyUtilities"
sys.path.append(path_package)
#import utility_revit_api as util_ra
import RevitUtilities as util_ra

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
last_start_time = start
average_times = list()

# Get all MEP instances (BuiltInCategories + MEP Categories)
all_instances = util_ra.get_BOQ_elements(rvt_doc)

# Start a blank table
table = list()

# Workset table for lookup
ws_table = rvt_doc.GetWorksetTable()

# Iterate over all
for i,elem in enumerate(all_instances):
    # Iterate, show status every N 
    N = 100
    if i % N == 0:
        this_loop_time = time.time()
        elapsed_time = this_loop_time - last_start_time
        last_start_time = this_loop_time
        
        # Add this to running average
        average_times.append(elapsed_time)

        # Keep the average to N
        N_avg = 5
        average_times = average_times[-N_avg:] 
        avg_elapsed_time = sum(average_times)/N_avg
        
        # Calcualate remaining time = Number Intervals * Time per Interval
        remaining_time_secs = (len(all_instances) - i)/N * avg_elapsed_time
        remaining_time_mins = remaining_time_secs/60
        
        print("{} of {} over {} seconds, remaining: {} minutes".format(i,len(all_instances),
                                                                       elapsed_time,remaining_time_mins))
        #print(average_times)
    
    # Break loop during testing
    if i >= 5000 and 1:
        break
    
    # Start a dict for this row of the table
    this_row = dict()
    
    # Get data for PRIMARY KEY
    # -Category
    try:
        this_row["Category"] = elem.Category.Name
    except:
        print(elem)
        print(elem.Name)
        print(elem.Category)
        #continue
        raise
    
    # -Workset    
    try:
        this_ws = ws_table.GetWorkset(elem.WorksetId)
        this_row["Workset"] = this_ws.Name
    except:
        #print("Workset problem",elem)
        this_row["Workset"] = "UNDEFINED"
        pass

    # -Family    
    try:
        this_row["Family"] = elem.Symbol.FamilyName
    except:
        #print("Family problem",elem)
        this_row["Family"] = "UNDEFINED"
        pass
    
    # -Type      
    try:
        this_row["Type"] = elem.Name
    except:
        #print("Type problem",elem)
        this_row["Type"] = "UNDEFINED"            
        pass
    
    this_row["Size"] = util_ra.get_parameter_plain_string(elem, "Size", flg_DNE=True)

    this_row["System"] = util_ra.get_parameter_plain_string(elem, "System Abbreviation", flg_DNE=True)

    # Get data for REPORTING
    this_row["Description"] = util_ra.get_parameter_plain_string(elem, "Description", flg_DNE=True)
    
    this_row["ifcDescription"] = util_ra.get_parameter_plain_string(elem, "ifcDescription", flg_DNE=True)
    
    this_row["Length"] = util_ra.get_parameter_value_str(elem, "Length", flg_DNE=True)
    
    this_row["Area"] = util_ra.get_parameter_value_str(elem, "Area", flg_DNE=True)
    
    this_row["Id"] = elem.Id
    
    this_row["IKEA Item Code"] = util_ra.get_parameter_plain_string(elem, "IKEA Item Code", flg_DNE=True)
    
    this_row["IKEA Cost Group"] = util_ra.get_parameter_plain_string(elem, "IKEA Cost Group", flg_DNE=True)
    
    # Enforce strings, over all elements in dict row
    #---TODO: Is this still necessary? 
    for k in this_row:
        try:
            this_row[k] = this_row[k].encode('ascii','ignore')
        except:
            #logging.error("Problem encoding {} = {} {} of elem {}".format(k,this_row[k], type(this_row[k]), elem.Name))
            #print(sys.exc_info()[0])
            continue

    # Append and end loop
    table.append(this_row)

#print(util_ra.format_as_table(table, table[0].keys(),table[0].keys()))

end = time.time()
        
logging.info("Built table of {} elements over {} seconds".format(len(table),end - start))
    
out_path = r"C:\CesCloud Revit\_03_IKEA_Working_Folder\BOQ Development\BOQ ALL5.csv"
with open(out_path,'wb') as csv_file:
    writer = csv.DictWriter(csv_file, delimiter=';', fieldnames=table[0].keys())
    writer.writeheader()
    writer.writerows(table)
    
end = time.time()
        
logging.info("Wrote table to {} over {} seconds".format(out_path,end - start))
    
logging.info("---DONE---".format())

