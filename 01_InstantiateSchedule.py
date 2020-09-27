"""
01 AutoColumnSchedule - Instantiate schedule family
"""
__author__ = 'Nicholas Eaw'
__version__ = '1.0.0'
__date created__ = '27/09/2020'

#load the Python Standard and DesignScript Libraries
import sys
import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = DocumentManager.Instance.CurrentDBDocument

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

import msvcrt

clr.AddReference("RevitAPIUI")
from  Autodesk.Revit.UI import *

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

#inputs
schedule = UnwrapElement(IN[0])
startpoint = Point.Create(XYZ(0,0,0)).Coord
OUTPUT = []
pNames = ["Number of Column", "Number of Row"]

#get all levels in the project
levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()

#get all views in the project
allViews = FilteredElementCollector(doc).OfClass(View).ToElements()

#get only legend view
view = []
for i in allViews:
	if i.ViewType == ViewType.Legend:
		if i.Name == "Column Schedule":
			view.append(i)

#get structural columns in the project
columns = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType().ToElements()

#get mark from columns
mark = []
for i in columns:
	p = i.LookupParameter("Mark").AsString()
	mark.append(p)

#remove null from list
mark = list(filter(None, mark))

#retrieve unique column marks from list
columnmarks = list(set(mark))

#set list of values for column and row arrays
nrows = len(levels) - 1
ncolumns = len(columnmarks)
pValues = [ncolumns, nrows]

#start transaction
TransactionManager.Instance.EnsureInTransaction(doc)

#create column schedule
cs = doc.Create.NewFamilyInstance(startpoint, schedule, view[0])

#loop through list of parameter names and values
for n,v in zip(pNames, pValues):
	p = cs.LookupParameter(n)
	p.Set(v)

#end transaction
TransactionManager.Instance.TransactionTaskDone()

#wrap
OUTPUT.append(cs.ToDSType(True))

#task dialog box
button = TaskDialogCommonButtons.None
result = TaskDialogResult.Ok
message = "Column schedule is created according to the number of levels and column types in the model. There is " + nrows.ToString() + " levels and " + ncolumns.ToString() + " types of columns."  
msgbox = TaskDialog.Show('Column Schedule', message, button)

#Assign your output to the OUT variable.
OUT = msgbox
