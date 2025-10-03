'[path=\Projects\Project A\Diagram Group]
'[group=Diagram Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' Script Name: Add Plateau Tags
' Author: Geert Bellekens
' Purpose: Add plateau tags for each element on this diagram
' Date: 2025-05-13

const outPutName = "Add Plateau Tags"
const templateLegendGUID = "{64C42112-AED2-4f9a-AEA0-85779EBE18F5}"

sub main
	'reset output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName

	dim diagram as EA.Diagram
	set diagram = Repository.GetCurrentDiagram
	if diagram is nothing then
		exit sub
	end if
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName & " for '"& diagram.Name &"'", 0
	'do the actual work
	addPlateauTags diagram
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName & " for '"& diagram.Name &"'", 0
	
end sub

function addPlateauTags(diagram)
	'first save diagram
	Repository.SaveDiagram diagram.DiagramID
	'get plateau
	dim plateau as EA.Element
	set plateau = nothing
	if diagram.ParentID > 0 then
		set plateau = Repository.GetElementByID(diagram.ParentID)
		if not plateau is nothing then
			if not plateau.Stereotype = "ArchiMate_Plateau" then
				set plateau = nothing
			end if
		end if
	end if
	if plateau is nothing then
		'let the user select a plateau
		dim plateauID
		plateauID = Repository.InvokeConstructPicker("IncludedTypes=Class;StereoType=ArchiMate_Plateau;") 
		if plateauID > 0 then
			set plateau = Repository.GetElementByID(plateauID)
		end if
	end if
	if plateau is nothing then
		Repository.WriteOutput outPutName, now() & " No plateau selected", 0
		exit function
	end if
	dim items
	set items = getSelectedElements(diagram)
	if not items.Count > 0 then
		set items = getAllItemsOnDiagram(diagram, plateau.Name)
	end if
	'add tagged values
	dim item
	for each item in items
		Repository.WriteOutput outPutName, now() & " Adding tag to '" & item.Name &"'", 0
		dim tag as EA.TaggedValue
		set tag = item.TaggedValues.AddNew(plateau.Name, "Unchanged")
		tag.Notes = "Values: New,Change,Delete,Unchanged" & vbNewLine & "Default: Unchanged"
		tag.update
	next
	'add the legend
	addPlateauLegend diagram, plateau
	'reload diagram
	Repository.ReloadDiagram diagram.DiagramID
end function

function addPlateauLegend(diagram, plateau)
	'check if legend is already on the diagram
	if isLegendOnDiagram(diagram) then
		exit function 'already on diagram, no need to continue
	end if
	'get the legend from the template package
	dim templateLegend as EA.Element
	set templateLegend = Repository.GetElementByGuid(templateLegendGUID)
	if templateLegend is nothing then
		Repository.WriteOutput outPutName, now() & " ERROR: could not find template legend with ea_guid: '" & templateLegendGUID & "'", 0
		exit function
	end if
	'create new legend
	dim legend as EA.Element
	dim diagramParent as EA.Element
	set diagramParent = getDiagramParent(diagram)
	set legend = diagramParent.Elements.AddNew(templateLegend.Name,"Text")
	legend.Subtype = templateLegend.Subtype
	legend.StyleEx = replace( templateLegend.StyleEx, "TaggedValue.Plateau Name", "TaggedValue." & plateau.Name)
	legend.Update
	'create t_xref record, copying properties from template legend
	dim sqlUpdate
	sqlUpdate = "insert into t_xref                                                                                         " & vbNewLine & _
				" select newID(), x.name, x.type, x.Visibility, x.NameSpace, x.Requirement, x.[Constraint]                  " & vbNewLine & _
				" , x.Behavior, x.Partition, x.Description, '" & legend.ElementGUID & "' , x.supplier, x.link               " & vbNewLine & _
				" from t_xref x                                                                                             " & vbNewLine & _
				" where x.client = '" & templateLegend.ElementGUID & "'                                                     " & vbNewLine & _
				" and x.Name = 'CustomProperties'                                                                           "
	Repository.Execute sqlUpdate
	'add legend to diagram
	addElementToDiagram legend, diagram, 100, 100
end function

function getDiagramParent(diagram)
	dim parent as EA.Element
	set parent = nothing
	if diagram.ParentID > 0 then
		set parent = Repository.GetElementByID(diagram.ParentID)
	end if
	if parent is nothing then
		set parent = Repository.GetPackageByID(diagram.PackageID)
	end if
	'return 
	set getDiagramParent = parent
end function


function isLegendOnDiagram(diagram)
	dim found
	found = false
	dim sqlGetData
	sqlGetdata = "select o.Object_ID from t_diagramobjects dod            " & vbNewLine & _
				" inner join t_object o on o.Object_ID = dod.Object_ID   " & vbNewLine & _
				" 		and o.Object_Type = 'Text'                       " & vbNewLine & _
				" 		and o.NType = '76'                               " & vbNewLine & _
				" where dod.Diagram_ID = " & diagram.DiagramID
	dim result
	set result = getArrayListFromQuery(sqlGetData)
	if result.Count > 0 then
		found = true
	end if
	'return
	isLegendOnDiagram = found
end function

function getAllItemsOnDiagram(diagram, plateauName)
	dim items
	set items = getAllElementsOnDiagram(diagram, plateauName)
	dim connectors
	set connectors = getAllConnectorsOnDiagram(diagram, plateauName)
	'add connectors to items
	items.AddRange(connectors)
	'return
	set getAllItemsOnDiagram = items
end function

function getAllConnectorsOnDiagram(diagram, plateauName)
	dim sqlGetData
	sqlGetData = "select dl.ConnectorID from t_diagramlinks dl                  " & vbNewLine & _
				" inner join t_diagram d on d.Diagram_ID = dl.DiagramID        " & vbNewLine & _
				" where d.ea_guid = '" & diagram.DiagramGUID & "'              " & vbNewLine & _
				" and dl.Hidden = 0                                            " & vbNewLine & _
				" and not exists	                                           " & vbNewLine & _
				" 	(select tv.ea_guid from t_connectortag tv                  " & vbNewLine & _
				" 	where tv.ElementID = dl.ConnectorID                        " & vbNewLine & _
				" 	and tv.Property = '" & plateauName & "')                   "			
	dim connectors
	set connectors = getConnectorsFromQuery(sqlGetData)
	'return
	set getAllConnectorsOnDiagram = connectors
end function

function getAllElementsOnDiagram(diagram, plateauName)
	dim sqlGetData
	sqlGetData = "select dod.Object_ID from t_diagramobjects dod          " & vbNewLine & _
				" inner join t_object o on o.Object_ID = dod.Object_ID   " & vbNewLine & _
				" 				and o.Stereotype is not null             " & vbNewLine & _
				" where dod.Diagram_ID = " & diagram.DiagramID & "       " & vbNewLine & _
				" and not exists	                                     " & vbNewLine & _
				" 	(select tv.ea_guid from t_objectproperties tv        " & vbNewLine & _
				" 	where tv.Object_ID = o.Object_ID                     " & vbNewLine & _
				" 	and tv.Property = '" & plateauName & "')                        "
	dim allElements
	set allElements = getElementsFromQuery(sqlGetData)
	'return
	set getAllElementsOnDiagram = allElements
end function

function getSelectedElements(diagram)
	dim selectedElements
	set selectedElements = CreateObject("System.Collections.ArrayList")
	dim selectedDiagramObjects
	set selectedDiagramObjects = diagram.SelectedObjects
	dim selectedDiagramObject as EA.DiagramObject
	for each selectedDiagramObject in selectedDiagramObjects
		dim selectedElement
		set selectedElement = Repository.GetElementByID(selectedDiagramObject.ElementID)
		if not selectedElement is nothing then
			selectedElements.Add selectedElement
		end if
	next	
	set getSelectedElements = selectedElements
end function

main
'test
'function test
'		'reset output tab
'	Repository.CreateOutputTab outPutName
'	Repository.ClearOutput outPutName
'	Repository.EnsureOutputVisible outPutName
'
'	dim diagram as EA.Diagram
'	set diagram = Repository.GetDiagramByGuid("{BC928698-ADB9-4ca1-82B9-F507613CEF9C}")
'	if diagram is nothing then
'		exit function
'	end if
'	'set timestamp
'	Repository.WriteOutput outPutName, now() & " Starting " & outPutName & " for '"& diagram.Name &"'", 0
'	'do the actual work
'	addPlateauTags diagram
'	'set timestamp
'	Repository.WriteOutput outPutName, now() & " Finished " & outPutName & " for '"& diagram.Name &"'", 0
'end function