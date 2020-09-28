"""
03 AutoColumnSchedule - Set column mark and size and create tags

__author__ = 'Nicholas Eaw'
__version__ = '1.0.0'
__date created__ = '27/09/2020'
"""

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

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

import System 

#inputs
output = []
outlist = []
category = UnwrapElement(IN[0])
tag1 = UnwrapElement(IN[1])
tag2 = UnwrapElement(IN[2])
tag3 = UnwrapElement(IN[3])
tag4 = UnwrapElement(IN[4])
drop = UnwrapElement(IN[5])

bic = System.Enum.ToObject(BuiltInCategory, category.Id.IntegerValue)

#start transaction
TransactionManager.Instance.EnsureInTransaction(doc)

#retrieve schedule instantiated
try:
	cschedule = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType().ToElements()
	for i in cschedule:
		if i.Name == "Column Schedule v01":
			outlist.append(i)

#if not successful, output fail message
except Exception, e:
	outlist.append("Failed\n" + e.message)
	
#end transaction	
TransactionManager.Instance.TransactionTaskDone()

#get parameter values
pNames = ["Column Schedule Width","Column Schedule Height", "Head Column Width", "Head Row Width", "Column Width", "Row Height", "Cell Row Height"]
pValues = []
for n in pNames:
	p = outlist[-1].LookupParameter(n).AsValueString()
	pValues.append(p)

#convert to int
pValues = [int(i) for i in pValues]

#convert from mm to ft
pValuesConverted = []
for i in pValues:
	pValuesConverted.append(i/304.8)
	
#create dictionary
cdict = dict(zip(pNames, pValuesConverted)) 

#get parameter values for column array
pColumnArray = int(outlist[-1].LookupParameter("Number of Column").AsValueString())
pRowArray = int(outlist[-1].LookupParameter("Number of Row").AsValueString())

#start position for detail item
xpos = -(cdict['Column Schedule Width']/2) + cdict['Head Column Width'] + cdict['Column Width']/2
ypos = -(cdict['Column Schedule Height']/2) + cdict['Head Row Width'] + (cdict['Row Height'] - cdict['Cell Row Height'] * 3)/2 + cdict['Cell Row Height'] * 3

#start position for annotations
axpos = -(cdict['Column Schedule Width']/2) + cdict['Head Column Width'] + (cdict['Column Width'] - (700/304.8))/2 + 700/304.8
aypos = -(cdict['Column Schedule Height']/2) + cdict['Head Row Width'] + cdict['Cell Row Height'] * 2.5

#difference
axposdiff = xpos - axpos
ayposdiff = ypos - aypos

#start position for column mark
cxpos = xpos 
cypos = -(cdict['Column Schedule Height']/2) + cdict['Head Row Width']/2

#create array of x points for column mark
cxposArray = [cxpos]
for i in range(pColumnArray - 1):
	cxpos = cxpos + cdict['Column Width']
	cxposArray.append(cxpos)

#get all views in the project
allViews = FilteredElementCollector(doc).OfClass(View).ToElements()

#get only legend view
view = []
for i in allViews:
	if i.ViewType == ViewType.Legend:
		if i.Name == "Column Schedule":
			view.append(i)

#get all levels in the project
levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()

#sort levels according to elevations
elevations = []
for l in levels:
	elevations.append(l.Elevation)

sortedlevels = [j for i,j in sorted(zip(elevations,levels))]

#get all the columns for each level
columnsLevel = []
rcColumns = []

for i in range(len(sortedlevels)):
	levelsId = sortedlevels[i].Id
	filter = ElementLevelFilter(levelsId)
	columnsLevel.append(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WherePasses(filter).ToElements())

#filter out steel columns
filteredColumnList = [[i for i in nested if str(i.StructuralMaterialType) == "Concrete"] for nested in columnsLevel]

#get column size
columnsize = []
for i in range(len(filteredColumnList)):
	columnsize.append([])
	for j in range(len(filteredColumnList[i])):
		columnTypeId = filteredColumnList[i][j].GetTypeId()
		columnElement = doc.GetElement(columnTypeId)
		dia = columnElement.LookupParameter("b")
		if dia:
			columnsize[i].append(dia.AsValueString())

#convert string of column sizes to int
columnsize = [list(map(int, i)) for i in columnsize]

#filter out empty lists
columnsize = [x for x in columnsize if x]

#get column marks
columnmark = []
for i in range(len(filteredColumnList)):
	columnmark.append([])
	for j in range(len(filteredColumnList[i])):
		p = filteredColumnList[i][j].LookupParameter("Mark").AsString()
		columnmark[i].append(p)

#filter out empty lists
columnmark = [x for x in columnmark if x]

#get unique sizes and mark
combine = zip(columnmark, columnsize)
combinelist = []
for n,v in combine:
	combinelist.append(zip(n,v))

columnrepository = []
for x in combinelist:
	uniquelist = sorted(set(tuple(x)))
	columnrepository.append(uniquelist)

#retrieve column marks for scheduling
columnmarkschedule = []
for i in range(len(columnrepository)):
	for j in range(len(columnrepository[i])):
		columnmarkschedule.append(columnrepository[i][j][0])

columnmarkschedule = sorted(list(set(columnmarkschedule)))

newlist = []
for i in range(len(columnrepository)):
	newlist.append([])
	addnew = pColumnArray - len(columnrepository[i])
	for j in range(addnew):
		newlist[i].insert(j,[])

#get column mark index
markindex = []
for i in columnmarkschedule:
	markindex.append(columnmarkschedule.index(i))

#get index of column mark to insert
insert = []
for i in range(len(columnrepository)):
	insert.append([])
	for j in range(len(columnrepository[i])):
		idx = columnmarkschedule.index(columnrepository[i][j][0])
		insert[i].append(idx)

#insert into new list
for i in range(len(columnrepository)):
	for j in range(len(columnrepository[i])):
		k = insert[i][j]
		newlist[i].insert(k, columnrepository[i][j])

#transpose list
tnewlist = list(zip(*newlist))

#start transaction
TransactionManager.Instance.EnsureInTransaction(doc)

#collect elements in view that need to be tagged
elementCollector = FilteredElementCollector(doc, view[0].Id).OfCategory(bic)
#collect tags present
tagCollector = FilteredElementCollector(doc, view[0].Id).OfClass(IndependentTag)

#check elements not tagged
notTagged = []
for e in elementCollector:
	tagged = False
	for t in tagCollector:
		if t.TaggedLocalElementId == e.Id:
			tagged = True
	if not tagged:
		if e.Name != "Column Schedule v01" and e.Name != "Main Bar Array_Y":
			notTagged.append(e)

#create tags for elements
for e in notTagged:
	x1 = e.Location.Point.X - axposdiff
	y1 = e.Location.Point.Y - ayposdiff
	eloc1 = Point.Create(XYZ(x1,y1,0)).Coord	
	t1 = IndependentTag.Create(doc, view[0].Id, Reference(e), False, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, eloc1)
	t1.ChangeTypeId(tag1.Id)
	output.append(t1)

#create 2nd tag
elementCollector = FilteredElementCollector(doc, view[0].Id).OfCategory(bic)
#collect tags present
tagCollector = FilteredElementCollector(doc, view[0].Id).OfClass(IndependentTag)

#check elements tagged
Tagged = []
for e in elementCollector:
	for t in tagCollector:
		if t.TaggedLocalElementId == e.Id:
			tagged = True
	if tagged:
		if e.Name != "Column Schedule v01" and e.Name != "Main Bar Array_Y":
			Tagged.append(e)

#create 2nd and 3rd tag
for e in Tagged:
	x2 = e.Location.Point.X - axposdiff 
	y2 = e.Location.Point.Y - ayposdiff - cdict['Cell Row Height']
	eloc2 = Point.Create(XYZ(x2,y2,0)).Coord	
	t2 = IndependentTag.Create(doc, view[0].Id, Reference(e), False, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, eloc2)
	t2.ChangeTypeId(tag2.Id)
	x3 = x2  
	y3 = y2 - cdict['Cell Row Height']	
	eloc3 = Point.Create(XYZ(x3,y3,0)).Coord
	t3 = IndependentTag.Create(doc, view[0].Id, Reference(e), False, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, eloc3)
	t3.ChangeTypeId(tag3.Id)

#create sublist for column marks
colsublist = []
counterlist = []

for i in range(len(tnewlist)):
	counter = 0
	for j in range(len(tnewlist[i])):
		if len(tnewlist[i][j]) == 0:
			counter = counter + 1
	counterlist.append(counter)

for i in range(0, len(Tagged), pRowArray-1):
	for num in counterlist:
		div = pRowArray - num
		break
	sublist = Tagged[i:i + div]
	colsublist.append(sublist)

#transpose list to get 1st row
transposeTaggedList = list(zip(*colsublist))

#create column mark
for i in range(len(colsublist)):
	x = cxposArray[i]
	y = cypos
	t4 = IndependentTag.Create(doc, view[0].Id, Reference(colsublist[i][0]), False, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, colsublist[i][0].Location.Point)
	t4.ChangeTypeId(tag4.Id)
	newpoint = Point.Create(XYZ(x,y,0)).Coord
	t4.Location.Move(newpoint - t4.TagHeadPosition)
	
#end transaction	
TransactionManager.Instance.TransactionTaskDone()

#collect all textnote types in project
textNotes = FilteredElementCollector(doc).OfClass(TextNoteType).WhereElementIsElementType().ToElements()

#get 3mm type
textOutput = []
for i in textNotes:
	n = i.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
	if n == "3mm":
		textOutput.append(i)

#get start position for level text notes
lxpos = -(cdict['Column Schedule Width']/2) + 100/304.8
lypos = -(cdict['Column Schedule Height']/2) + cdict['Head Row Width'] + 200/304.8

#get start position for drop symbol
dsxpos = lxpos
dsypos = -(cdict['Column Schedule Height']/2) + cdict['Head Row Width']

#create array of y points for drop symbols
dsyposArray = [dsypos]
for i in range(len(sortedlevels)):
	dsypos = dsypos + cdict['Row Height']
	dsyposArray.append(dsypos)

#create array of y points for text notes
lyposArray = [lypos]
for i in range(pRowArray):
	lypos = lypos + cdict['Row Height']
	lyposArray.append(lypos)

#start transaction
TransactionManager.Instance.EnsureInTransaction(doc)

#create textnote
for i in range(len(sortedlevels)):
	x = lxpos
	y = lyposArray[i]
	textNotePoint = Point.Create(XYZ(x,y,0)).Coord
	textNote = TextNote.Create(doc, doc.ActiveView.Id, textNotePoint, sortedlevels[i].Name.upper(), textOutput[0].Id)
	y1 = dsyposArray[i]
	dropSymbolPoint = Point.Create(XYZ(x,y1,0)).Coord
	dp = doc.Create.NewFamilyInstance(dropSymbolPoint, drop, view[0])

#end transaction
TransactionManager.Instance.TransactionTaskDone()

clean = []
for i in range(len(tnewlist)):
	clean.append([])
	for j in range(len(tnewlist[i])):
		if len(tnewlist[i][j]) != 0:
			clean[i].append(tnewlist[i][j])
			
#outputs
OUT = colsublist, clean
