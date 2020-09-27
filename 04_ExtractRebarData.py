"""
04 AutoColumnSchedule - Extract data from Excel and load into column schedule
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

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

import System 

import math

#inputs
output = []
outlist = []
category = UnwrapElement(IN[0])
data = IN[1]

bic = System.Enum.ToObject(BuiltInCategory, category.Id.IntegerValue)

#transpose data
tdata = list(map(list,zip(*data)))

#get rebar details only
filtertdata = tdata[2:]

#filter out empty items in list
rebardetails = []
for i in range(len(filtertdata)):
	rebardetails.append([])
	for j in range(len(filtertdata[i])):
		if filtertdata[i][j]:
			rebardetails[i].append(filtertdata[i][j])
#get rebar
rebar = []
for i in range(len(rebardetails)):
	rebar.append([])
	for j in range(0, len(rebardetails[i]), 4):
		rebar[i].append(rebardetails[i][j])

rebar_cleaned = []
for i in range(len(rebar)):
	c = rebar[i][:-1]
	rebar_cleaned.append(c)

rebar_cleaned = [x[::-1] for x in rebar_cleaned]

#get column sizes
csize = []
for i in range(len(rebardetails)):
	csize.append([])
	for j in range(3, len(rebardetails[i]), 4):
		csize[i].append(rebardetails[i][j])

csize = [x[::-1] for x in csize]

#get bar type
bartype = []
for i in range(len(rebar_cleaned)):
	bartype.append([])
	for j in range(len(rebar_cleaned[i])):
		b = rebar_cleaned[i][j][2:3]
		bartype[i].append(b)

#get rebar diameter
rebardia = []
for i in range(len(rebar_cleaned)):
	rebardia.append([])
	for j in range(len(rebar_cleaned[i])):
		dia = float(rebar_cleaned[i][j][3:5])
		rebardia[i].append(dia)

#get number of rebars
rebarno = []
for i in range(len(rebar_cleaned)):
	rebarno.append([])
	for j in range(len(rebar_cleaned[i])):
		dia = int(rebar_cleaned[i][j][0:2])
		rebarno[i].append(dia)

#get rebar area
rebararea = []
for i in range(len(rebardia)):
	rebararea.append([])
	for j in range(len(rebardia[i])):
		area = math.pi * (rebardia[i][j]/2)**2
		rebararea[i].append(area)		

#get rebar area required
rebararearequired = []
for i in range(len(rebararea)):
	rebararearequired.append([])
	for j in range(len(rebararea[i])):
		required = rebararea[i][j] * rebarno[i][j] 
		rebararearequired[i].append(required)
		
#get rebar percentage required
rebarpercentage = []
for i in range(len(rebararearequired)):
	rebarpercentage.append([])
	for j in range(len(rebararearequired[i])):
		required = rebararearequired[i][j] / 0.01 / (math.pi * (csize[i][j]*0.5)**2)
		rebarpercentage[i].append(required)

#get links
links = []
for i in range(len(rebardetails)):
	links.append([])
	for j in range(1, len(rebardetails[i]), 4):
		links[i].append(rebardetails[i][j])

links = [x[::-1] for x in links]

#get link bar type
linkbartype = []
for i in range(len(links)):
	linkbartype.append([])
	for j in range(len(links[i])):
		b = links[i][j][0:1]
		linkbartype[i].append(b)

#get link bar size
linkbarsize = []
for i in range(len(links)):
	linkbarsize.append([])
	for j in range(len(links[i])):
		b = links[i][j][1:3]
		linkbarsize[i].append(int(b))

#get link bar spacing
linkbarspacing = []
for i in range(len(links)):
	linkbarspacing.append([])
	for j in range(len(links[i])):
		b = links[i][j][4:7]
		linkbarspacing[i].append(int(b))
		
#get hooks
hooks = []
for i in range(len(rebardetails)):
	hooks.append([])
	for j in range(2, len(rebardetails[i]), 4):
		hooks[i].append(rebardetails[i][j])

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

#get number of levels with columns
columnlevellist = []
for i in columnsLevel:
	if len(i) != 0:
		columnlevellist.append(len(i))

levelno = len(columnlevellist)

#get number of column marks
colmarkno = len(columnmarkschedule)

#get all views in the project
allViews = FilteredElementCollector(doc).OfClass(View).ToElements()

#get only legend view
view = []
for i in allViews:
	if i.ViewType == ViewType.Legend:
		if i.Name == "Column Schedule":
			view.append(i)
			
#collect elements in view that need to be tagged
elementCollector = FilteredElementCollector(doc, view[0].Id).OfCategory(bic)

for e in elementCollector:
	if e.Name != "Column Schedule v01" and e.Name != "Main Bar Array_Y":
		outlist.append(e)

columnrepositoryreversed = list(reversed(columnrepository))
#create sublist for column marks
colsublist = []
counterlist = []

for i in range(len(columnrepositoryreversed)):
	counter = len(columnrepositoryreversed[i])
	counterlist.append(counter)

for i in range(0, len(outlist), levelno-1):
	for num in counterlist:
		div = num
		break
	sublist = outlist[i:i + div]
	colsublist.append(sublist)
	
#create ref array
refArray = []
ref = []
for item in outlist:
	refplane = item.GetReferences(FamilyInstanceReferenceType.WeakReference)
	ref.append(refplane)
	
for i in ref:
	rArr = ReferenceArray()
	for j in i:
		rArr.Append(j)
	refArray.append(rArr)		

#get center points of column detail items
centerpoints = []
for i in outlist:
	centerpoints.append(i.Location.Point)

#get column size
size = []
for i in outlist:
	p = int(i.LookupParameter("Column Size").AsValueString()) / 304.8
	size.append(p)

#get starting point of line
xdimstartArray = []
ydimstartArray = []

for i in range(len(centerpoints)):
	xdimstartpoint = centerpoints[i].X - (size[i]/2)
	xdimstartArray.append(xdimstartpoint)
	ydimstartpoint = centerpoints[i].Y + (size[i]/2) + 150/304.8
	ydimstartArray.append(ydimstartpoint)
	
#get ending point of line
xdimendArray = []
ydimendArray = []

for i in range(len(centerpoints)):
	xdimendpoint = centerpoints[i].X + (size[i]/2)
	xdimendArray.append(xdimendpoint)
	ydimendpoint = centerpoints[i].Y + (size[i]/2) + 150/304.8
	ydimendArray.append(ydimendpoint)

#start transaction
TransactionManager.Instance.EnsureInTransaction(doc)

#create lines
lines = []
for i in range(len(outlist)):
	startpoint = Point.Create(XYZ(xdimstartArray[i], ydimstartArray[i],0)).Coord
	endpoint = Point.Create(XYZ(xdimendArray[i], ydimendArray[i],0)).Coord
	line = Line.CreateBound(startpoint, endpoint)
	lines.append(line)
	
#create dimensions
for i in range(len(lines)):
	dim = doc.Create.NewDimension(view[0], lines[i], refArray[i])

#end transaction
TransactionManager.Instance.TransactionTaskDone()

#outputs
OUT = colsublist, rebarpercentage, rebardia, bartype, linkbarsize, linkbarspacing, linkbartype
