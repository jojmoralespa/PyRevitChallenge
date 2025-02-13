"""Numbering the selected elements according the relative grid position and position from main element selected"""

__title__ = 'Grid-Based\nNumbering'

from pyrevit import forms
from Autodesk.Revit.DB import ExternalDefinitionCreationOptions, SpecTypeId
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import Selection,TaskDialog,TaskDialogCommonButtons
import clr


app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
rvt_year = int(app.VersionNumber)

#Main section

grids = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType())

listOfCoordinates = []

for i in range(len(grids)):
    for j in range(i + 1, len(grids)):

        curve1 = grids[i].Curve
        curve2 = grids[j].Curve
        intersectionResults = clr.Reference[IntersectionResultArray]()
        intersectionBetween = curve1.Intersect(curve2, intersectionResults)
        
        if intersectionBetween == SetComparisonResult.Overlap:
            direction1 = curve1.Direction 
            direction2 = curve2.Direction
            gridSquare = ""
            
            interResult = intersectionResults.Value
            intersectionPoint = interResult[0].XYZPoint
            if abs(direction1.X) > abs(direction1.Y) and abs(direction2.Y) > abs(direction2.X): 
                gridSquare = "{}-{}".format(grids[j].Name,grids[i].Name)
            else:
                gridSquare = "{}-{}".format(grids[i].Name,grids[j].Name)

            nameAndCoordinatePair = {
                "name":gridSquare,
                "coordinate": intersectionPoint
            }

            listOfCoordinates.append(nameAndCoordinatePair)
        else:
            continue
                
def getOrCreateProjectParameter(parameterName, buitInCategory):
    sharedParamsFile = app.OpenSharedParameterFile()
    if not sharedParamsFile:
        forms.alert("No shared parameter file found. Please create one.", exitscript=True)
        return None
    
    group = None
    for g in sharedParamsFile.Groups:
        if g.Name == "Custom Parameters":
            group = g
            break

    if not group:
        group = sharedParamsFile.Groups.Create("Custom Parameters")
    
    definition = None
    for group in sharedParamsFile.Groups:
        for defn in group.Definitions:
            if defn.Name == parameterName:
                definition = defn
                break

    if not definition:
        opt = ExternalDefinitionCreationOptions(parameterName, SpecTypeId.String.Text)
        definition = group.Definitions.Create(opt)

    categorySet = app.Create.NewCategorySet()
    categoryView = Category.GetCategory(doc,buitInCategory)
    categorySet.Insert(categoryView)

    newInstanceBinding = app.Create.NewInstanceBinding(categorySet)

    if rvt_year >= 2024:
        parameterGroup = GroupTypeId.AnalysisResults
    else:
        parameterGroup = BuiltInParameterGroup.PG_ANALYSIS_RESULTS

    t = Transaction(doc, "Add Parameter: " + parameterName)
    t.Start()

    try:
        doc.ParameterBindings.Insert(definition, newInstanceBinding, parameterGroup)
    except:
        forms.alert('Failed to Add: {}'.format(definition.Name))
    t.Commit()
    
def select_elements(coordList):
    try:
        TaskDialog.Show("Grid-Based Numbering - Selection of elements","Select Elements to modify ( all of the same category)",TaskDialogCommonButtons.Ok)
        selection = uidoc.Selection.PickObjects(Selection.ObjectType.Element, "select elements")
        selectedElements = [doc.GetElement(ref.ElementId) for ref in selection]

        TaskDialog.Show("Grid-Based Numbering - Selection of elements", "Select the reference element", TaskDialogCommonButtons.Ok)
        refElement = uidoc.Selection.PickObject(Selection.ObjectType.Element, "Select the first element for numbering")
        refElement = doc.GetElement(refElement.ElementId)

        selectedElements.sort(key=lambda e: refElement.Location.Point.DistanceTo(e.Location.Point))

        buitInCategory = selectedElements[0].Category.BuiltInCategory

        getOrCreateProjectParameter("Grid Square",buitInCategory)
        getOrCreateProjectParameter("Number", buitInCategory)

        t = Transaction(doc, "Update Parameters")
        t.Start()

        for index, elem in enumerate(selectedElements, start=1):
            parameterGrid = elem.LookupParameter("Grid Square")
            parameterNumber = elem.LookupParameter("Number")

            if parameterGrid and parameterGrid.IsReadOnly == False:
                closest_grid = min(coordList, key=lambda c: elem.Location.Point.DistanceTo(c["coordinate"]))
                parameterGrid.Set(closest_grid["name"])

            if parameterNumber and parameterNumber.IsReadOnly == False:
                parameterNumber.Set(str(index))

        t.Commit()

    except:
        forms.alert("the elements cannot be selected.")
    
select_elements(listOfCoordinates)