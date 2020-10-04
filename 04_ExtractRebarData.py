"""
04 AutoColumnSchedule - Extract data from Excel and load into column schedule

__author__ = 'Nicholas Eaw'
__version__ = '1.1.0'
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

import math

#inputs
output = []
outlist = []
category = UnwrapElement(IN[0])
data = IN[1]

bic = System.Enum.ToObject(BuiltInCategory, category.Id.IntegerValue)

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
		if "Round" in columnElement.FamilyName.ToString(): 
			dia = columnElement.LookupParameter("b")
			columnsize[i].append(dia.AsValueString())
		elif "Rectangular" in columnElement.FamilyName.ToString():
			columnsize[i].append([])
			w1 = columnElement.LookupParameter("b")
			w2 = columnElement.LookupParameter("h")
			columnsize[i][j].append(w1.AsValueString())
			columnsize[i][j].append(w2.AsValueString())


#convert string of column sizes to int
def convert_to_int(lists):
	return [int(el) if not isinstance(el,list) else convert_to_int(el) for el in lists]
	
columnsize = convert_to_int(columnsize)

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

col = []
bool = []
newlst = []
uniquelst = []
col1 = []

def get_unique(lists):
	for floor in lists:
		for column in floor:
			for i in column:
				if not isinstance(i ,list):
					t = False
				elif isinstance(i, list):
					t = True
		bool.append(t)
	
	for i in range(len(bool)):
		if not bool[i]:
			unique = sorted(set(tuple(lists[i])))
			uniquelst.append(unique)
		elif bool[i]:
			uniquelst.append([])
			for j in range(len(lists[i])):
				uniquelst[i].append([])
				uniquelst[i][j].append(lists[i][j][0])
				for k in range(len(lists[i][j][1])):
					uniquelst[i][j].append(lists[i][j][1][k])
					
			unique = sorted(set(map(tuple, uniquelst[i])))
			col.append(unique)
	
	for i in range(len(bool)):
		if bool[i]:
			uniquelst[i] = col[i-1]
	
	uniquelist = map(list, uniquelst)
	
	return uniquelist

columnrepository = get_unique(combinelist)

#retrieve column marks for scheduling
columnmarkschedule = []
for i in range(len(columnrepository)):
	for j in range(len(columnrepository[i])):
		columnmarkschedule.append(columnrepository[i][j][0])

columnmarkschedule = sorted(list(set(columnmarkschedule)))

newlist = []
for i in range(len(columnrepository)):
	newlist.append([])
	addnew = len(columnmarkschedule) - len(columnrepository[i])
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

#filter out empty lists
clist = []
for i in range(len(tnewlist)):
	clist.append([])
	for j in range(len(tnewlist[i])):
		if len(tnewlist[i][j]) != 0:
			clist[i].append(tnewlist[i][j])
			
newclist = []			
for i in range(len(clist)):
	newclist.append([])
	for j in range(len(clist[i])):
		newclist[i].append(clist[i][j][1:])

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

#get column type
ctype = []
for i in range(len(rebardetails)):
	ctype.append([])
	for j in range(3, len(rebardetails[i]), 4):
		ctype[i].append(rebardetails[i][j])

ctype = [x[::-1] for x in ctype]

bartype = []
bartypeindex = []

#get index of bar type
for i in range(len(rebar_cleaned)):
	bartypeindex.append([])
	for j in range(len(rebar_cleaned[i])):
		for pos,char in enumerate(rebar_cleaned[i][j]):
			if char == 'T':
				bartypeindex[i].append(int(pos))

#get bar type				
for i in range(len(rebar_cleaned)):
	bartype.append([])
	for j in range(len(rebar_cleaned[i])):
		b = rebar_cleaned[i][j][bartypeindex[i][j]:bartypeindex[i][j]+1]
		bartype[i].append(b)

#get rebar diameter
rebardia = []
for i in range(len(rebar_cleaned)):
	rebardia.append([])
	for j in range(len(rebar_cleaned[i])):
		dia = float(rebar_cleaned[i][j][bartypeindex[i][j]+1:bartypeindex[i][j]+3])
		rebardia[i].append(dia)

#get number of rebars
rebarno = []
for i in range(len(rebar_cleaned)):
	rebarno.append([])
	for j in range(len(rebar_cleaned[i])):
		dia = int(rebar_cleaned[i][j][0:bartypeindex[i][j]])
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
		if ctype[i][j] == "CIRCULAR":
			required = rebararearequired[i][j] / 0.01 / (math.pi * (newclist[i][j][0]*0.5)**2)
			rebarpercentage[i].append(required)
		else:
			required = rebararearequired[i][j] / 0.01 / (newclist[i][j][0] * newclist[i][j][1])
			rebarpercentage[i].append(required)

#get links
links = []
for i in range(len(rebardetails)):
	links.append([])
	for j in range(1, len(rebardetails[i]), 4):
		links[i].append(rebardetails[i][j])

links = [x[::-1] for x in links]

linkbartypeindex = []
#get index of link bar type
for i in range(len(links)):
	linkbartypeindex.append([])
	for j in range(len(links[i])):
		for pos,char in enumerate(links[i][j]):
			if char == 'T':
				linkbartypeindex[i].append(int(pos))
				
#get link bar type
linkbartype = []
for i in range(len(links)):
	linkbartype.append([])
	for j in range(len(links[i])):
		b = links[i][j][linkbartypeindex[i][j]:linkbartypeindex[i][j]+1]
		linkbartype[i].append(b)

#get link bar size
linkbarsize = []
for i in range(len(links)):
	linkbarsize.append([])
	for j in range(len(links[i])):
		b = links[i][j][linkbartypeindex[i][j]+1:linkbartypeindex[i][j]+3]
		linkbarsize[i].append(int(b))

#get link bar spacing
linkbarspacing = []
for i in range(len(links)):
	linkbarspacing.append([])
	for j in range(len(links[i])):
		b = links[i][j][linkbartypeindex[i][j]+4:linkbartypeindex[i][j]+7]
		linkbarspacing[i].append(int(b))
		
#get hooks
hooks = []
for i in range(len(rebardetails)):
	hooks.append([])
	for j in range(2, len(rebardetails[i]), 4):
		hooks[i].append(rebardetails[i][j])
OUT = hooks
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
circlerefArray = []
rect1refArray = []
rect2refArray = []
ref = []

for i in range(len(outlist)):
	ref.append([])
	refplane = outlist[i].GetReferences(FamilyInstanceReferenceType.WeakReference)
	ref[i].append(refplane)

width1List = []
width2List = []
circleList = []

#separate into circle and widths array
for i in range(len(outlist)):
	width1List.append([])
	width2List.append([])
	circleList.append([])
	for j in range(len(ref[i])):
		for k in range(len(ref[i][j])):
			if "Width1" in outlist[i].GetReferenceName(ref[i][j][k]):
				width1List[i].append(ref[i][j][k])
			elif "Width2" in outlist[i].GetReferenceName(ref[i][j][k]):
				width2List[i].append(ref[i][j][k])
			elif "Cir" in outlist[i].GetReferenceName(ref[i][j][k]):
				circleList[i].append(ref[i][j][k])

#create ref array for circular dimension
for i in circleList:
	rArr = ReferenceArray()
	for j in i:
		rArr.Append(j)
	circlerefArray.append(rArr)
	
#create ref array for rect dimension width 1
for i in width1List:	
	rArr = ReferenceArray()
	for j in i:
		rArr.Append(j)
	rect1refArray.append(rArr)

#create ref array for rect dimension width 2
for i in width2List:	
	rArr = ReferenceArray()
	for j in i:
		rArr.Append(j)
	rect2refArray.append(rArr)

#get center points of column detail items
centerpoints = []
for i in range(len(colsublist)):
	centerpoints.append([])
	for j in range(len(colsublist[i])):
		centerpoints[i].append(colsublist[i][j].Location.Point)

#get column size
colsize = []
for i in range(len(colsublist)):
	colsize.append([])
	for j in range(len(colsublist[i])):
		if colsublist[i][j].Name == "Link_Main":
			p = int(colsublist[i][j].LookupParameter("Column Size").AsValueString()) / 304.8
			colsize[i].append(p)
		else:
			p1 = int(colsublist[i][j].LookupParameter("Width 1").AsValueString()) / 304.8
			p2 = int(colsublist[i][j].LookupParameter("Width 2").AsValueString()) / 304.8
			colsize[i].append([p1, p2])

#get starting point of line
xdimstartArray = []
ydimstartArray = []

for i in range(len(centerpoints)):
	xdimstartArray.append([])
	ydimstartArray.append([])
	for j in range(len(centerpoints[i])):
		if colsublist[i][j].Name == "Link_Main":
			xdimstartpoint = centerpoints[i][j].X - (colsize[i][j]/2)
			xdimstartArray[i].append(xdimstartpoint)
			ydimstartpoint = centerpoints[i][j].Y + (colsize[i][j]/2) + 150/304.8
			ydimstartArray[i].append(ydimstartpoint)
		else:
			x1dimstartpoint = centerpoints[i][j].X - (colsize[i][j][0]/2)
			x2dimstartpoint = centerpoints[i][j].X + (colsize[i][j][1]/2) + 150/304.8
			xdimstartArray[i].append([x1dimstartpoint, x2dimstartpoint])
			y1dimstartpoint = centerpoints[i][j].Y + (colsize[i][j][1]/2) + 150/304.8
			y2dimstartpoint = centerpoints[i][j].Y - (colsize[i][j][1]/2) 
			ydimstartArray[i].append([y1dimstartpoint, y2dimstartpoint])
		
#get ending point of line
xdimendArray = []
ydimendArray = []

for i in range(len(centerpoints)):
	xdimendArray.append([])
	ydimendArray.append([])
	for j in range(len(centerpoints[i])):
		if colsublist[i][j].Name == "Link_Main":
			xdimendpoint = centerpoints[i][j].X + (colsize[i][j]/2)
			xdimendArray[i].append(xdimendpoint)
			ydimendpoint = centerpoints[i][j].Y + (colsize[i][j]/2) + 150/304.8
			ydimendArray[i].append(ydimendpoint)
		else:
			x1dimendpoint = centerpoints[i][j].X + (colsize[i][j][0]/2) 
			x2dimendpoint = centerpoints[i][j].X + (colsize[i][j][1]/2) + 150/304.8
			xdimendArray[i].append([x1dimendpoint, x2dimendpoint])
			y1dimendpoint = centerpoints[i][j].Y + (colsize[i][j][1]/2) + 150/304.8
			y2dimendpoint = centerpoints[i][j].Y + (colsize[i][j][1]/2) 
			ydimendArray[i].append([y1dimendpoint, y2dimendpoint])

#start transaction
TransactionManager.Instance.EnsureInTransaction(doc)

#create lines
lines = []
for i in range(len(colsublist)):
	for j in range(len(colsublist[i])):
		if colsublist[i][j].Name == "Link_Main":
			startpoint = Point.Create(XYZ(xdimstartArray[i][j], ydimstartArray[i][j],0)).Coord
			endpoint = Point.Create(XYZ(xdimendArray[i][j], ydimendArray[i][j],0)).Coord
			line = Line.CreateBound(startpoint, endpoint)
#			detailline = doc.Create.NewDetailCurve(view[0], line)
			lines.append(line)
		else:
			startpoint1 = Point.Create(XYZ(xdimstartArray[i][j][0], ydimstartArray[i][j][0],0)).Coord
			endpoint1 = Point.Create(XYZ(xdimendArray[i][j][0], ydimendArray[i][j][0],0)).Coord
			startpoint2 = Point.Create(XYZ(xdimstartArray[i][j][1], ydimstartArray[i][j][1],0)).Coord
			endpoint2 = Point.Create(XYZ(xdimendArray[i][j][1], ydimendArray[i][j][1],0)).Coord
			line1 = Line.CreateBound(startpoint1, endpoint1)
#			detailline1 = doc.Create.NewDetailCurve(view[0], line1)
			line2 = Line.CreateBound(startpoint2, endpoint2)
#			detailline2 = doc.Create.NewDetailCurve(view[0], line2)
			lines.append([line1, line2])
			
	
#create dimensions
for i in range(len(outlist)):
	if outlist[i].Name == "Link_Main":
		dim = doc.Create.NewDimension(view[0], lines[i], circlerefArray[i])
	else:
		dim1 = doc.Create.NewDimension(view[0], lines[i][0], rect1refArray[i])
		dim2 = doc.Create.NewDimension(view[0], lines[i][1], rect2refArray[i])

#end transaction
TransactionManager.Instance.TransactionTaskDone()

#outputs
OUT = colsublist, rebarpercentage, rebardia, bartype, linkbarsize, linkbarspacing, linkbartype
