"""
02 AutoColumnSchedule - Create and place column detail items 

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

#inputs
circolumndetail = UnwrapElement(IN[0])
sqcolumndetail = UnwrapElement(IN[1])
outlist = []
OUTPUT = []

#retrieve schedule instantiated
try:
	cschedule = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType().ToElements()
	for i in cschedule:
		if i.Name == "Column Schedule v01":
			outlist.append(i)

#if not successful, output fail message
except Exception, e:
	outlist.append("Failed\n" + e.message)
	
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

#create array of x points
xposArray = [xpos]
for i in range(pColumnArray - 1):
	xpos = xpos + cdict['Column Width']
	xposArray.append(xpos)

#create array of y points
yposArray = [ypos]
for i in range(pRowArray - 1):
	ypos = ypos + cdict['Row Height']
	yposArray.append(ypos)
	
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

#check if columns are circular or rectangular
colfamily = []
for i in range(len(filteredColumnList)):
	colfamily.append([])
	for j in range(len(filteredColumnList[i])):
		columnTypeId = filteredColumnList[i][j].GetTypeId()
		columnElement = doc.GetElement(columnTypeId)
		colfamily[i].append(columnElement.FamilyName.ToString())		

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

#get unique column mark and type
combinemarktype = zip(columnmark, colfamily)
colmarktype = []
for n,v in combinemarktype:
	colmarktype.append(zip(n,v))

colmarktyperepo = []
for x in colmarktype:
	uniquelist = sorted(set(tuple(x)))
	colmarktyperepo.append(uniquelist)

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
	addnew = pColumnArray - len(columnrepository[i])
	for j in range(addnew):
		newlist[i].insert(j,[])
		
newlist1 = []
for i in range(len(columnrepository)):
	newlist1.append([])
	addnew = pColumnArray - len(columnrepository[i])
	for j in range(addnew):
		newlist1[i].insert(j,[])
		
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

for i in range(len(colmarktyperepo)):
	for j in range(len(colmarktyperepo[i])):
		k = insert[i][j]
		newlist1[i].insert(k, colmarktyperepo[i][j])

#transpose list
tnewlist = list(zip(*newlist))
tnewlist1 = list(zip(*newlist1))

#create line starting and ending points for x and y
xlinestartpoint = -(cdict['Column Schedule Width']/2) + cdict['Head Column Width']

ylinestartpoint = -(cdict['Column Schedule Height']/2) + cdict['Head Row Width'] + (cdict['Cell Row Height'] * 3)

xlineendpoint = -(cdict['Column Schedule Width']/2) + cdict['Head Column Width'] + cdict['Column Width']

ylineendpoint = -(cdict['Column Schedule Height']/2) + cdict['Head Row Width'] + cdict['Row Height']

#create array
xArrayLineStartPoints = [xlinestartpoint]
for i in range(pColumnArray - 1):
	xlinestartpoint = xlinestartpoint + cdict['Column Width']
	xArrayLineStartPoints.append(xlinestartpoint)
	
yArrayLineStartPoints = [ylinestartpoint]
for i in range(pRowArray - 1):
	ylinestartpoint = ylinestartpoint + cdict['Row Height']
	yArrayLineStartPoints.append(ylinestartpoint)
	
xArrayLineEndPoints = [xlineendpoint]
for i in range(pColumnArray - 1):
	xlineendpoint = xlineendpoint + cdict['Column Width']
	xArrayLineEndPoints.append(xlineendpoint)

yArrayLineEndPoints = [ylineendpoint]
for i in range(pRowArray - 1):
	ylineendpoint = ylineendpoint + cdict['Row Height']
	yArrayLineEndPoints.append(ylineendpoint)

#start transaction
TransactionManager.Instance.EnsureInTransaction(doc)

#create array of points and place column detail item 
for i in range(len(tnewlist)):
	for j in range(len(tnewlist[i])):
		if len(tnewlist[i][j]) == 0:
			linestartpoint = Point.Create(XYZ(xArrayLineStartPoints[i], yArrayLineStartPoints[j],0)).Coord
			lineendpoint = Point.Create(XYZ(xArrayLineEndPoints[i], yArrayLineEndPoints[j],0)).Coord
			line = Line.CreateBound(linestartpoint, lineendpoint)
			detailline = doc.Create.NewDetailCurve(view[0], line)		
		else:
			position = Point.Create(XYZ(xposArray[i],yposArray[j],0)).Coord
			if "Round" in tnewlist1[i][j][1]:
				cd = doc.Create.NewFamilyInstance(position, circolumndetail, view[0])
			elif "Rectangular" in tnewlist1[i][j][1]:
				cd = doc.Create.NewFamilyInstance(position, sqcolumndetail, view[0])

#end transaction
TransactionManager.Instance.TransactionTaskDone()


# Assign your output to the OUT variable.
OUT = "Success"
