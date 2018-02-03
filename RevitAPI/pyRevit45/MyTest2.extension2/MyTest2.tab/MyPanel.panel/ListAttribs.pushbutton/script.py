from scriptutils import logger, this_script
from scriptutils.userinput import CommandSwitchWindow
from revitutils import doc, selection

#from collections import namedtuple
#from customcollections import DefaultOrderedDict

out = this_script.output
#ParamDef = namedtuple('ParamDef', ['name', 'type'])

# noinspection PyUnresolvedReferences
from Autodesk.Revit.DB import Element, CurveElement, ElementId, \
                              StorageType, ParameterType

__doc__ = 'Testing' \

__title__ = 'List Attributes'

# Iterate elements
for el in selection.elements:
    print("Element {}".format(el))
    #print(dir(el))
    for param in el.ParametersMap:
        pdef = param.Definition
        print("\t{} - {} = {} as {}".format(pdef.ParameterType,
                                       pdef.Name,param.AsValueString(),param.StorageType))