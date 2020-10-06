"""
05 AutoColumnSchedule - Create rebar hooks for rectangular columns

__author__ = 'Nicholas Eaw'
__version__ = '1.0.0'
__date created__ = '06/10/2020'
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
hookdetail = UnwrapElement(IN[2])
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

#get column type
ctype = []
for i in range(len(rebardetails)):
	ctype.append([])
	for j in range(3, len(rebardetails[i]), 4):
		ctype[i].append(rebardetails[i][j])
		
#get hooks
hooks = []
for i in range(len(rebardetails)):
	hooks.append([])
	for j in range(2, len(rebardetails[i]), 4):
		hooks[i].append(rebardetails[i][j])

#filter out for rectangular columns
recthooks = []
for i in range(len(hooks)):
	recthooks.append([])
	for j in range(len(hooks[i])):
		if ctype[i][j] == "RECTANGULAR":
			recthooks[i].append(hooks[i][j])
			
#filter out empty lists	
recthooks = [i for i in recthooks if i]

hookbartypeindex = []

#get index of hook bar type
for i in range(len(hooks)):
	hookbartypeindex.append([])
	for j in range(len(hooks[i])):
		for pos,char in enumerate(hooks[i][j]):
			if 'T' in char:
				hookbartypeindex[i].append(int(pos))
				
#filter out empty lists
hookbartypeindex = [i for i in hookbartypeindex if i]

#get link bar type
hookbartype = []
for i in range(len(recthooks)):
	hookbartype.append([])
	for j in range(len(recthooks[i])):
		b = recthooks[i][j][hookbartypeindex[i][j]:hookbartypeindex[i][j]+1]
		hookbartype[i].append(b)

#get number of hooks
numhooks = []
for i in range(len(recthooks)):
	numhooks.append([])
	for j in range(len(recthooks[i])):
		n = recthooks[i][j][0:hookbartypeindex[i][j]]
		numhooks[i].append(int(n))

#get hook rebar size
hookrebar = []
for i in range(len(recthooks)):
	hookrebar.append([])
	for j in range(len(recthooks[i])):
		n = recthooks[i][j][hookbartypeindex[i][j]+1:hookbartypeindex[i][j]+3]
		hookrebar[i].append(int(n))

#get hook spacing
hookspacing = []
for i in range(len(recthooks)):
	hookspacing.append([])
	for j in range(len(recthooks[i])):
		n = recthooks[i][j][hookbartypeindex[i][j]+4:hookbartypeindex[i][j]+7]
		hookspacing[i].append(int(n))

#flatten
hookspacingf = [item for sublist in hookspacing for item in sublist]

#get all views in the project
allViews = FilteredElementCollector(doc).OfClass(View).ToElements()

#get only legend view
view = []
for i in allViews:
	if i.ViewType == ViewType.Legend:
		if i.Name == "Column Schedule":
			view.append(i)
			
#collect detail items
elementCollector = FilteredElementCollector(doc, view[0].Id).OfCategory(bic)

for e in elementCollector:
	if e.Name != "Column Schedule v01" and e.Name != "Main Bar Array_Y":
		outlist.append(e)

#if rectangular detail items present, filter them out
rectdetailitems = []
for item in outlist:
	if item.Name == "Column Rebar By %":
		rectdetailitems.append(item)

#retrieve data from detail items
bar1List = []
bar2List = []
width1List = []
width2List = []
coverList = []
rebardiaList = []
linkdiaList = []

for item in rectdetailitems:
	bar1 = item.LookupParameter("No of Bar 1").AsValueString()
	width1 = item.LookupParameter("Width 1").AsValueString()
	bar2 = item.LookupParameter("No of Bar 2").AsValueString()
	width2 = item.LookupParameter("Width 2").AsValueString()
	cover = item.LookupParameter("Cover").AsValueString()
	rebardia = item.LookupParameter("Bar Diameter").AsValueString()
	linkdia = item.LookupParameter("Main Link Diameter").AsValueString()
	bar1List.append(int(bar1))
	bar2List.append(int(bar2))
	width1List.append(int(width1))
	width2List.append(int(width2))
	coverList.append(int(cover))
	rebardiaList.append(int(rebardia))
	linkdiaList.append(int(linkdia))

#get number of hooks for width 1 & 2
hooks1 = []
hooks2 = []
bar1spacing = []
bar2spacing = []


for i in range(len(width1List)):
	spacing1 = (width1List[i] - (coverList[i]*2 + linkdiaList[i]*2 + rebardiaList[i]))/ (bar1List[i]-1)
	hookno = (width1List[i] - (coverList[i]*2 + linkdiaList[i]*2 + rebardiaList[i] + spacing1*2))/ spacing1 + 1
	bar1spacing.append(spacing1)
	n = 1
	while hookspacingf[i] * n < spacing1 * hookno:
		n +=1
		if hookspacingf[i] * n > spacing1 * hookno:
			hooks1.append(n)
			break			

for i in range(len(width2List)):
	spacing2 = (width2List[i] - (coverList[i]*2 + linkdiaList[i]*2 + rebardiaList[i]))/ (bar2List[i]-1)
	hookno = (width2List[i] - (coverList[i]*2 + linkdiaList[i]*2 + rebardiaList[i] + spacing2*2))/ spacing2 + 1
	bar2spacing.append(spacing2)
	n = 1
	while hookspacingf[i] * n < spacing2 * hookno:
		n +=1
		if hookspacingf[i] * n > spacing2 * hookno:
			hooks2.append(n)
			break
			
#convert from mm to ft
width1List = [i/304.8 for i in width1List]
width2List = [i/304.8 for i in width2List]
coverList = [i/304.8 for i in coverList]
linkdiaList = [i/304.8 for i in linkdiaList]
rebardiaList = [i/304.8 for i in rebardiaList]
bar1spacing = [i/304.8 for i in bar1spacing]
bar2spacing = [i/304.8 for i in bar2spacing]

#get center points of column detail items
startpoints = []

for item in rectdetailitems:
	pos = item.Location.Point
	startpoints.append(pos)
	
#get start points for hook in width 1	
xstartpoint1 = []
ystartpoint1 = []

for i in range(len(rectdetailitems)):
	xpos = startpoints[i].X - width1List[i]/2 + coverList[i] + linkdiaList[i] + rebardiaList[i]/2 + bar1spacing[i]
	xstartpoint1.append(xpos)
	ypos = startpoints[i].Y 
	ystartpoint1.append(ypos)	

#create array points for width 1
xstartpointArray1 = []
ystartpointArray1 = []

for i in range(len(rectdetailitems)):
	xstartpointArray1.append([])
	for j in range(hooks1[i]):
		if j == 0:
			xnpoint = xstartpoint1[i] 
			xstartpointArray1[i].append(xnpoint)
		elif j > 0:
			xnpoint = xnpoint + 2*bar1spacing[i]
			xstartpointArray1[i].append(xnpoint)

for i in range(len(rectdetailitems)):
	ystartpointArray1.append([])
	for j in range(hooks1[i]):
		ynpoint = ystartpoint1[i] 
		ystartpointArray1[i].append(ynpoint)

#get start points for hook in width 2
xstartpoint2 = []
ystartpoint2 = []

for i in range(len(rectdetailitems)):
	xpos = startpoints[i].X
	xstartpoint2.append(xpos)
	ypos = startpoints[i].Y - width2List[i]/2 + coverList[i] + linkdiaList[i] + rebardiaList[i]/2 + bar2spacing[i]
	ystartpoint2.append(ypos)
	
#create array points for width 2
xstartpointArray2 = []
ystartpointArray2 = []

for i in range(len(rectdetailitems)):
	xstartpointArray2.append([])
	for j in range(hooks2[i]):
		xnpoint = xstartpoint2[i] 
		xstartpointArray2[i].append(xnpoint)

for i in range(len(rectdetailitems)):
	ystartpointArray2.append([])
	for j in range(hooks2[i]):
		if j == 0:
			ynpoint = ystartpoint2[i] 
			ystartpointArray2[i].append(ynpoint)
		elif j > 0:
			ynpoint = ynpoint + 2*bar2spacing[i]
			ystartpointArray2[i].append(ynpoint)			

#create line of axis to rotate hook for width 2 side
axis = []
for i in range(len(xstartpointArray2)):
	axis.append([])
	for j in range(len(xstartpointArray2[i])):
		p1 = Point.Create(XYZ(xstartpointArray2[i][j] , ystartpointArray2[i][j],0)).Coord
		p2 = Point.Create(XYZ(xstartpointArray2[i][j] , ystartpointArray2[i][j], 10)).Coord
		line = Line.CreateBound(p1, p2)
		axis[i].append(line)
	
#start transaction
TransactionManager.Instance.EnsureInTransaction(doc)

hd1List = []
hd2List = []

#create hook detail items for width 1
for i in range(len(xstartpointArray1)):
	hd1List.append([])
	for j in range(len(xstartpointArray1[i])):	
		position1 = Point.Create(XYZ(xstartpointArray1[i][j], ystartpointArray1[i][j],0)).Coord
		hd1 = doc.Create.NewFamilyInstance(position1, hookdetail, view[0])
		hd1List[i].append(hd1)
		p1 = hd1.LookupParameter("L5")
		hd1_L5 = width2List[i] - coverList[i]*2
		p1.Set(hd1_L5)

#create hook detail items for width 2
for i in range(len(xstartpointArray2)):
	hd2List.append([])
	for j in range(len(xstartpointArray2[i])):
		position2 = Point.Create(XYZ(xstartpointArray2[i][j], ystartpointArray2[i][j],0)).Coord
		hd2 = doc.Create.NewFamilyInstance(position2, hookdetail, view[0])
		hd2List[i].append(hd2)
		ElementTransformUtils.RotateElement(doc, hd2.Id, axis[i][j], math.pi/2.0)
		p2 = hd2.LookupParameter("L5")
		hd2_L5 = width1List[i] - coverList[i]*2
		p2.Set(hd2_L5)

#end transaction
TransactionManager.Instance.TransactionTaskDone()

#get total number of hooks detail items created
totalhooks = []
for i in range(len(recthooks)):
	totalhooks.append([])
	for j in range(len(recthooks[i])):
		n = len(hd1List[i]) + len(hd2List[i])
		totalhooks[i].append(n)
	
#create text for hooks
hookstext = []
for i in range(len(recthooks)):
	hookstext.append([])
	for j in range(len(recthooks[i])):
		text = str(totalhooks[i][j]) + str(hookbartype[i][j]) + str(hookrebar[i][j]) + '-' + str(hookspacing[i][j])
		hookstext[i].append(text)

#collect all textnote types in project
textNotes = FilteredElementCollector(doc).OfClass(TextNoteType).WhereElementIsElementType().ToElements()

#get 3mm type
textOutput = []
for i in textNotes:
	n = i.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
	if n == "2.5mm":
		textOutput.append(i)

columnschedule = []
#start transaction
TransactionManager.Instance.EnsureInTransaction(doc)

#retrieve schedule instantiated
try:
	cschedule = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType().ToElements()
	for i in cschedule:
		if i.Name == "Column Schedule v01":
			columnschedule.append(i)

#if not successful, output fail message
except Exception, e:
	columnschedule.append("Failed\n" + e.message)
	
#end transaction	
TransactionManager.Instance.TransactionTaskDone()

#get parameter values
pNames = ["Column Schedule Width","Column Schedule Height", "Head Column Width", "Head Row Width", "Column Width", "Row Height", "Cell Row Height"]
pValues = []
for n in pNames:
	p = columnschedule[-1].LookupParameter(n).AsValueString()
	pValues.append(p)

#convert to int
pValues = [int(i) for i in pValues]

#convert from mm to ft
pValuesConverted = []
for i in pValues:
	pValuesConverted.append(i/304.8)
	
#create dictionary
cdict = dict(zip(pNames, pValuesConverted)) 

#start transaction
TransactionManager.Instance.EnsureInTransaction(doc)

hookstext = [item for sublist in hookstext for item in sublist]

#create textnote
for i in range(len(hookstext)):
	xpos = rectdetailitems[i].Location.Point.X + 250/304.8
	ypos = rectdetailitems[i].Location.Point.Y - cdict['Row Height']/2 - cdict['Cell Row Height']/2 - 25/304.8
	position = Point.Create(XYZ(xpos,ypos,0)).Coord
	textNote = TextNote.Create(doc, doc.ActiveView.Id, position, hookstext[i], textOutput[0].Id)

#end transaction
TransactionManager.Instance.TransactionTaskDone()

#outputs
OUT = "Success"
