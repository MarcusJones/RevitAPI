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

#path_package = r"C:\Dropbox\00 CAD Standards\70 Revit Python\Packages\xlrd"
#sys.path.append(path_package)
#import xlrd

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

#-Paths---

#import clr
#clr.AddReference('ProtoGeometry')
#clr.AddReference('PresentationCore')
#clr.AddReference('RevitAPI')
#clr.AddReference('RevitAPIUI')
#clr.AddReference('System.Xml.Linq')
#from Autodesk.DesignScript.Geometry import *
#from Autodesk.Revit.UI import *
#from Autodesk.Revit.Attributes import *
#from System.Diagnostics import Process
#from Autodesk.Revit.DB import *
#from Autodesk.Revit.DB.Architecture import *
#from Autodesk.Revit.DB.Analysis import *

#clr.AddReference("RevitServices")
#import RevitServices 
#from RevitServices.Persistence import DocumentManager

#doc = DocumentManager.Instance.CurrentDBDocument
#uiapp = DocumentManager.Instance.CurrentUIApplication
#app = uiapp.Application

#import sys

grids = util_ra.get_grids(rvt_doc)

# Define the bounding grid lines
grid_bounds = {
    'part1' : {'left' : '1', 'right' : '7', 'top' : 'L', 'bot' : 'G'},
    'part2' : {'left' : '6', 'right' : '12', 'top' : 'L', 'bot' : 'G'},
    'part3' : {'left' : '11', 'right' : '17', 'top' : 'L', 'bot' : 'G'},
    'part4' : {'left' : '1', 'right' : '7', 'top' : 'G', 'bot' : 'A'},
    'part5' : {'left' : '6', 'right' : '12', 'top' : 'G', 'bot' : 'A'},
    'part6' : {'left' : '11', 'right' : '17', 'top' : 'G', 'bot' : 'A'},
    }

# Assign the bound names to the grid elements directly
for part in grid_bounds:
    boundlist = list()
    for bound in grid_bounds[part]:
        this_grid_name = grid_bounds[part][bound]
        boundlist.append(this_grid_name)
        #print(part, bound, grid_bounds[part][bound])
        grid_bounds[part][bound] = grids[this_grid_name]
    logging.info("Part view named {} bounded by grid set {}.".format(part,boundlist))

# Create the boxes, with scaling factor, assign it to part definition
for part in grid_bounds:
    bound_box = util_ra.get_bound_box(grid_bounds[part],0.1)
    grid_bounds[part] = bound_box


#print(grid_bounds)
#raise
#logging.info("Part view named {} bounded by grid set {}.".format(part,boundlist))

# flattend_out = []
# 
# for k in grid_bounds:
#     print(k)
#     flattend_out.append([k, grid_bounds[k]])
#     

logging.info("{} dependant view definitions created".format(len(grid_bounds)))

this_view = rvt_doc.ActiveView
logging.info("This view: {}".format(this_view))

for part_name in grid_bounds:
    curve_loop = grid_bounds[part_name]
    new_view = util_ra.create_dependent(rvt_doc,this_view, part_name)
    util_ra.apply_crop(rvt_doc, new_view, curve_loop)
    #logging.debug("*** Processed new view {}".format(new_view))
    #crop = CurveLoop()
    
#     for line in curve_loop:
#         # Need to APPEND the lines
#         crop.Append(line)
# 
# raise
# for this_def in def_list:
#     print(this_def)
#     name = this_def[0]
# 
#     crop = CurveLoop()
#     for line in this_def[1]:
#         # Need to APPEND the lines
#         crop.Append(line)
# 
# 
#     #print("Name ", name)
#     #print("Crop1 ", type(crop), crop)
#     #print("Crop2 ", type(crop[0]), crop[0])
#     #box = this_def[0]
#     #print(this_def)
#     print("Creating {} over {} crop region".format(name, crop))
#     #print(name)
#     #print(type(this_def))
#     new_view = create_dependent(this_view, name)
#     apply_crop(rvt_doc,new_view, crop)
# 
# 
# #for part in grid_bounds:
# #    part_def = grid_bounds[part]
# #    print("Creating {}".format(part))
# #    create_dependent(this_view, part, part_def)



#----
logging.info("---DONE---".format())


