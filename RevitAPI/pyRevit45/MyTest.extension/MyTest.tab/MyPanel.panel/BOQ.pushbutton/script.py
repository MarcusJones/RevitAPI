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

#===============================================================================
#--- MAIN SCRIPT
#===============================================================================
print("Doc", doc)
print("uidoc", uidoc)

print("Get sorted elements")

sorted_elems = util_elems.get_sort_all_elements(doc)
for k in sorted_elems:
    break
    print(k, len(sorted_elems[k]))

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

first_wall = sorted_elems["Walls"][0]
print(first_wall)
util_params.table_parameters(first_wall)

util_params.table_parameters_sameline(first_wall)


print("{:20}".format("-name-").encode('utf-8'), end="")
print("{:20}".format("-ParameterGroup-").encode('ascii'), end="")
print("{:30}".format("-ParameterType-").encode('ascii'), end="")
print("{:30}".format("-Value String-").encode('ascii'), end="")
print("{:30}".format("-String-").encode('ascii'), end="")
print("{:30}".format("-UnitType-").encode('ascii'), end="")
print("")


print("{:20}".format("-name-"), end="")
print("{:20}".format("-ParameterGroup-"), end="")
print("{:30}".format("-ParameterType-"), end="")
print("{:30}".format("-Value String-"), end="")
print("{:30}".format("-String-"), end="")
print("{:30}".format("-UnitType-"), end="")
print("{} {}".format(1,23))
print("")

print("DONE")










