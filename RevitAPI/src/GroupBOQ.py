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
# Main script
#===============================================================================
logging.info("Python version : {}".format(sys.version))
#logging.info("uidoc : {}".format(rvt_uidoc))
#logging.info("doc : {}".format(rvt_doc))
#logging.info("app : {}".format(rvt_app))
 

def main():
    logging.debug("Start")
    
    path_csv = r"C:\CesCloud Revit\_03_IKEA_Working_Folder\BOQ Development\BOQ ALL3.csv"

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

    keyed_boq = dict()
    
    for row in table_dict:
        #print(row)
        this_key = (row['Workset'],
            row['Category'],
            row['Family'],
            row['Type'],
            row['System'],
            row['Size'])
        
        # Create this row if not encountered
        if this_key not in keyed_boq:
            keyed_boq[this_key] = dict()
        
        # Add to the QTY
        if 'QTY' not in keyed_boq[this_key]:
            keyed_boq[this_key]['QTY'] = 1
        else: 
            keyed_boq[this_key]['QTY'] = keyed_boq[this_key]['QTY'] + 1
            
        # Add to the Length
        if 'Length' not in keyed_boq[this_key]:
            keyed_boq[this_key]['Length'] = 0
        else: 
            if row['Length'] != 'DNE' and row['Length'] != None:
                if re.match("^\d+$", row['Length']):
                    keyed_boq[this_key]['Length'] = keyed_boq[this_key]['Length']+int(row['Length'])
        
        # Copy the whole row
        keyed_boq[this_key]['row'] = row
            
    logging.debug("Grouped into {} rows".format(len(keyed_boq)))
    
    i = 0
    boq_table = list()
    for k in keyed_boq:
        # Move QTY to ROW
        keyed_boq[k]['row']['QTY'] = keyed_boq[k]['QTY']
        keyed_boq[k].pop('QTY')
        # Move Length to ROW
        keyed_boq[k]['row']['Length'] = keyed_boq[k]['Length']
        keyed_boq[k].pop('Length')
        # Final table
        keyed_boq[k]['row'].pop('Id')
        boq_table.append(keyed_boq[k]['row'])
        
    out_path = r"C:\CesCloud Revit\_03_IKEA_Working_Folder\BOQ Development\BOQ COMBINED1.csv"
    
    fieldnames = [
                  'Workset',                   
                  'Category', 
                  'Family',   
                  'Type', 
                  'System', 
                  'Size', 
                  'Description', 
                  'Area', 
                  'ifcDescription', 
                  'IKEA Item Code', 
                  'IKEA Cost Group',
                  'QTY', 
                  'Length', 
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
