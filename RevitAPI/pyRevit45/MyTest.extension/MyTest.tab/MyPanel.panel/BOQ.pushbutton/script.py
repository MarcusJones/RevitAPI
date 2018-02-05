#===============================================================================
#--- SETUP Config
#===============================================================================
from __future__ import print_function

#===============================================================================
#--- SETUP Logging
#===============================================================================
from scriptutils import logger, this_script
from scriptutils.userinput import CommandSwitchWindow

logger.info('Test Log Level :ok_hand_sign:')

logger.warning('Test Log Level')

logger.critical('Test Log Level')

#===============================================================================
#--- SETUP Standard modules
#===============================================================================
import sys
from collections import namedtuple
from customcollections import DefaultOrderedDict

#===============================================================================
#--- SETUP Custom modules
#===============================================================================
path_exergy_util_dir = r"C:\LOCAL_REPO\py_ExergyUtilities"

sys.path.insert(0,path_exergy_util_dir)

print(sys.path)
from RevitUtilities import utility_general as util_gen
from RevitUtilities import utility_get_elements as util_elems
from RevitUtilities import utility_parameters as util_params

print(util_gen)
#===============================================================================
#--- SETUP Add parent module
#===============================================================================

from revitutils import doc, selection

#===============================================================================
#--- pyRevit setup
#===============================================================================
out = this_script.output
ParamDef = namedtuple('ParamDef', ['name', 'type'])

# noinspection PyUnresolvedReferences
from Autodesk.Revit.DB import Element, CurveElement, ElementId, \
                              StorageType, ParameterType

__doc__ = 'Testing BOQ' \

__title__ = 'BOQ'

from revitutils import doc, uidoc

#===============================================================================
#--- Functions
#===============================================================================
def test_get_elements():
    print("Get all elems")
    elems = util_elems.get_all_elements(doc)
    
    print("All elements as ", type(elems))
    print("Total number elements:", len(elems))
    
    print("First element:", elems[0], type(elems[0]))

def get_el_counts():
    
    def quant_print(cat, this_dict): 
        print("{} {}".format(cat, len(this_dict[cat])))

    quant_print("Walls", sorted_elems)
    quant_print("Floors" , sorted_elems)
    quant_print("Structural Framing", sorted_elems)
    quant_print("Generic Models", sorted_elems)
    quant_print("Curtain Systems", sorted_elems)
    quant_print("Doors", sorted_elems)
    quant_print("Columns", sorted_elems)
    quant_print("Roofs", sorted_elems)
    quant_print("Windows", sorted_elems)
    quant_print("Structural Columns", sorted_elems)
    quant_print("Levels", sorted_elems)

def print_el_table():
    pass

    #util_params.table_parameters_scirpted(first_wall,this_script.output.print_code)

def print_sorted_elems(sorted_elems):
    for k in sorted_elems:
        print(k, len(sorted_elems[k]))

def get_walls():
    el = sorted_elems["Floors"][0]
    
    print(el)
    print("Name {}".format(el.Name))
    
        
    #first_wall = sorted_elems["Walls"][0]
    
    this_dict = util_params.get_all_params(first_wall)
    #print(first_wall.Name())
    print(this_dict)
    
    print(dir(first_wall))
    
#util_params.print_all_params(first_wall)
    
    
#===============================================================================
#--- MAIN SCRIPT
#===============================================================================

print("Doc", doc)
print("uidoc", uidoc)


print("Get sorted elements")


sorted_elems = util_elems.get_sort_all_elements(doc)


boq_elems = util_elems.get_BOQ_elements(doc)

print(len(boq_elems))


#get_walls()

print("DONE")










