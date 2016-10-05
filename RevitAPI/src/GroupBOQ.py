"""
Given an exported BOQ CSV from Revit
In format with PRIMARY KEY columns and DATA columns
Merge items with same PRIMARY KEY by summing relevant DATA columns
"""

#===============================================================================
# Set up
#===============================================================================
# Standard:
from config.config import *
import re
import logging.config
import unittest

#from ExergyUtilities.utility_inspect import get_self
#from sympy.categories.baseclasses import Category
#from statsmodels.genmod.families.family import Family

from collections import defaultdict
import sys

import csv
#===============================================================================
# Other modules
#===============================================================================

#===============================================================================
# Logging
#===============================================================================
print(ABSOLUTE_LOGGING_PATH)
logging.config.fileConfig(ABSOLUTE_LOGGING_PATH)
myLogger = logging.getLogger()
myLogger.setLevel("DEBUG")

#===============================================================================
# Definitions
#===============================================================================
def group_table(table_dict):
    keyed_boq = dict()
    
    # Iterate all rows
    for row in table_dict:
        #print(row)
        
        # The primary key is the following tuple
        this_key = (row['Workset'],
            row['Category'],
            row['Family'],
            row['Type'],
            row['System'],
            row['Size'])
        
        # Create this key row if not yet encountered
        if this_key not in keyed_boq:
            keyed_boq[this_key] = dict()
        
        # QTY is a COUNTED parameter
        if 'QTY' not in keyed_boq[this_key]:
            keyed_boq[this_key]['QTY'] = 1
        else: 
            keyed_boq[this_key]['QTY'] = keyed_boq[this_key]['QTY'] + 1

        # Area is a SUMMED parameter
        if 'Area' not in keyed_boq[this_key]:
            # Extract the integers
            area_ints = [int(s) for s in row['Area'].split() if s.isdigit()]
            if len(area_ints) > 0:
                this_area = area_ints.pop(0)
                # Initialize
                keyed_boq[this_key]['Area'] = this_area
            else:
                # If there is not a nice number, ignore this parameter for this key
                keyed_boq[this_key]['Area'] = None
        else:
            # Add to the existing
            area_ints = [int(s) for s in row['Area'].split() if s.isdigit()]
            if len(area_ints) > 0:
                this_area = area_ints.pop(0)
                if row['Area'] != 'DNE' and row['Area'] != None:
                    keyed_boq[this_key]['Area'] = keyed_boq[this_key]['Area'] + this_area
        
        
        #if row['ifcDescription'] == "BV_DU Duct Tees 550 SA":
        #    print(row)
        #    print(keyed_boq[this_key])
        
        # Length is a SUMMED parameter
        if 'Length' not in keyed_boq[this_key]:
            # Check if the length is actually a number
            if re.match("^\d+$", row['Length']):
                # Initialize the length
                keyed_boq[this_key]['Length'] = int(row['Length'])
            else:
                # If there is not a nice number, ignore this parameter for this key
                keyed_boq[this_key]['Length'] = None
        else: 
            # Add to the existing
            if row['Length'] != 'DNE' and row['Length'] != None:
                if re.match("^\d+$", row['Length']):
                    keyed_boq[this_key]['Length'] = keyed_boq[this_key]['Length']+int(row['Length'])
                
        # Copy the whole row
        keyed_boq[this_key]['row'] = row
    
    logging.debug("Grouped into {} rows".format(len(keyed_boq)))
    return keyed_boq


#===============================================================================
# Main script
#===============================================================================
logging.info("Python version : {}".format(sys.version))
#logging.info("uidoc : {}".format(rvt_uidoc))
#logging.info("doc : {}".format(rvt_doc))
#logging.info("app : {}".format(rvt_app))

def main():
    logging.debug("Start")
    
    path_csv = r"C:\CesCloud Revit\_03_IKEA_Working_Folder\BOQ Development\BOQ ALL4.csv"

    table_dict = list()
    
    keyed_boq = defaultdict(list)
    
    with open(path_csv) as csvfile:
        reader = csv.reader(csvfile,delimiter=';')
        headers = next(reader)        
        
        reader = csv.DictReader(csvfile,fieldnames=headers,delimiter=';')
        
        for row in reader:
            op_A = False
            for k in row:
                found = op_A or bool(row[k]) 
                if found: 
                    break
            if not found:
                break
            table_dict.append(row)
            
    logging.debug("Loaded {} rows".format(len(table_dict)))

    keyed_boq = group_table(table_dict)
    
    # The COUNTED and SUMMED parameters are extra columns
    # Overwrite the orginal columns with the counted and summed values
    boq_table = list()
    for k in keyed_boq:
        # Move QTY to ROW
        keyed_boq[k]['row']['QTY'] = keyed_boq[k]['QTY']
        keyed_boq[k].pop('QTY')
        # Move Length to ROW
        keyed_boq[k]['row']['Length'] = keyed_boq[k]['Length']
        keyed_boq[k]['row']['Length [mm]'] = keyed_boq[k]['row'].pop('Length')
        keyed_boq[k].pop('Length')
        # Move Area to ROW
        keyed_boq[k]['row']['Area'] = keyed_boq[k]['Area']
        keyed_boq[k]['row']['Area [m2]'] = keyed_boq[k]['row'].pop('Area')
        keyed_boq[k].pop('Area')
        # Final table
        keyed_boq[k]['row'].pop('Id')
        # Append the adjusted row
        boq_table.append(keyed_boq[k]['row'])
        
    out_path = r"C:\CesCloud Revit\_03_IKEA_Working_Folder\BOQ Development\BOQ COMBINED6.csv"
    
    # Define the order of fields in the final table
    fieldnames = [
                  'Workset',                   
                  'Category', 
                  'Family',   
                  'Type', 
                  'System', 
                  'Size', 
                  'Description', 
                  'Area [m2]', 
                  'ifcDescription', 
                  'IKEA Item Code', 
                  'IKEA Cost Group',
                  'QTY', 
                  'Length [mm]', 
                  ]
    
    with open(out_path,'w') as csv_file:
        writer = csv.DictWriter(csv_file, 
                                delimiter=';', 
                                fieldnames=fieldnames,
                                lineterminator='\n',)
        writer.writeheader()
        writer.writerows(boq_table)
        
        
    logging.debug("Done")


if __name__ == '__main__':
    main()
