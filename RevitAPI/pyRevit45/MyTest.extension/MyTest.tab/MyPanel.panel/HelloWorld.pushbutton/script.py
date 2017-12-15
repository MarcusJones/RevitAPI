from scriptutils import logger, this_script
from scriptutils.userinput import CommandSwitchWindow
from revitutils import doc, selection

from collections import namedtuple
from customcollections import DefaultOrderedDict

out = this_script.output
ParamDef = namedtuple('ParamDef', ['name', 'type'])

# noinspection PyUnresolvedReferences
from Autodesk.Revit.DB import Element, CurveElement, ElementId, \
                              StorageType, ParameterType

__doc__ = 'Testing' \


__title__ = 'MY Test Hello'


print('Hello22'.format())

rvt_uidoc = __revit__.ActiveUIDocument


this_sel = rvt_uidoc.Selection
this_sel = selection
print(this_sel)
print(dir(this_sel))
#print(this_sel.ToString())
elIDS = rvt_uidoc.Selection.GetElementIds()
print(elIDS)
mech_eq_ID = elIDS[0]
print(mech_eq_ID)
print(type(mech_eq_ID))
mech_eq_el = doc.GetElement(mech_eq_ID)
#print(mech_eq_el)
#print(dir(mech_eq_el))
#for i in dir(mech_eq_el):
#    print(i)
print("*********")
#room_attr = mech_eq_el.Room
#print(room_attr)
#print(type(room_attr))
def get_room(phase, ele):
    phase = list(doc.Phases)[1]
    room = mech_eq_el.Room[phase]
    #print(room.Name.ToString())
    this_number = room.Number
    print(dir(room))
    print(room.Parameters)
    for param in room.Parameters:
        #print("{} - {}".format(param.Definition.Name,param.AsValue()))
        if param.Definition.Name == "Name":
            this_name = param.AsString()
        #print("{} - {}".format(param.Definition.Name,param.AsString()))
        #print()
    return(this_name,this_number)