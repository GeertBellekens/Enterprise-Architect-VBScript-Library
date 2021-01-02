'[path=\Projects\Project MZ\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Add elements to glossary
' Author: Geert Bellekens
' Purpose: Add the elements in this package to the glossary diagram
' Date: 2019-02-06
'
const outPutName = "Add to Glossary"


function Main ()
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	if not selectedPackage is Nothing then
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'inform user
		Repository.WriteOutput outPutName, now() & " Starting add to Glossary for package '" & selectedPackage.Name & "'" , 0
		'do the actual work
		addElementsToGlossary selectedPackage
		'inform user
		Repository.WriteOutput outPutName, now() & " Finished add to Glossary for package '" & selectedPackage.Name & "'" , 0
	end if
end function

function addElementsToGlossary(package)
	'check if the opened diagram is a glossary diagram
	dim glossaryDiagram as EA.Diagram
	'set glossaryDiagram  = Repository.GetDiagramByGuid("{EEB08636-4E86-499f-96D5-FD110608967C}")
	set glossaryDiagram  = Repository.GetCurrentDiagram()
	'check if this is a glossary diagram
	if not glossaryDiagram is nothing then
		if instr(glossaryDiagram.StyleEx, "MDGDgm=Glossary Item Lists::GlossaryItemList") > 0 then
			addElementsToDiagram package, glossaryDiagram
		else
			'create new glossary diagram
			addToNewGlossaryDiagram package
		end if
	else
		'create new glossary diagram
		addToNewGlossaryDiagram package
	end if
end function

function addToNewGlossaryDiagram(package)
	'ask user if he wants to create a new diagram
	dim response
	response = Msgbox("The current diagram is not a glossary diagram." & vbNewLine & "Would you like to create a new glossary diagram?" , vbYesNo+vbQuestion, "Create new Glossary Diagram?")
	if not response = vbYes then
		exit function
	end if
	'Let user select a package
	dim diagramPackage as EA.Package
	set diagramPackage = selectPackage()
	'Let the user enter a diagram name
	dim diagramName
	diagramName = InputBox("Please enter the name for the glossary diagram", "Diagram Name", "Glossary")
	if len(diagramName) = 0 then
		exit function
	end if
	'create the new diagram
	dim glossaryDiagram as EA.Diagram
	set glossaryDiagram = diagramPackage.Diagrams.AddNew(diagramName, "Glossary Item Lists::GlossaryItemList")
	glossaryDiagram.Update
	'add the elements to the diagram
	addElementsToDiagram package, glossaryDiagram
end function

function addElementsToDiagram(package, diagram)
	'get elementsToAdd
	dim elements
	set elements = getNamedElementsNotOnDiagram(package, diagram)
	'add each element to the diagram
	dim x
	x = 0
	dim y
	y = 0
	dim element as EA.Element
	for each element in elements
		Repository.WriteOutput outPutName, now() & " Adding element '" & element.Name & "'" , 0
		'move coordinates
		x = x + 40
		Y = y + 40
		'add element
		dim positionString
		positionString = "l=" & x & ";r=" & x + 90 & ";t=" & y & ";b=" & y + 50 & ";"
		dim diagramObject as EA.DiagramObject
		set diagramObject  = diagram.DiagramObjects.AddNew(positionString,"")
		diagramObject.ElementID = element.ElementID
		diagramObject.Update
	next
	'reload diagram
	Repository.ReloadDiagram diagram.DiagramID
	'open the diagram if not already open
	Repository.OpenDiagram diagram.DiagramID
end function

function getNamedElementsNotOnDiagram(package, diagram)
	'get package tree id
	dim packageTreeIDs
	packageTreeIDs = getPackageTreeIDString(package)
	'make SQL Suery
	dim sqlGetData
	sqlGetData = 	"select o.Object_ID from t_object o                " & vbNewLine & _
					" where o.Name is not null                         " & vbNewLine & _
					" and o.Package_ID in (" & packageTreeIDs & ")     " & vbNewLine & _
					" and not exists (                                 " & vbNewLine & _
					" 	select * from t_diagramobjects do              " & vbNewLine & _
					" 	where do.Object_ID = o.Object_ID               " & vbNewLine & _
					" 	and do.Diagram_ID = " & diagram.DiagramID & "  " & vbNewLine & _
					" 	)                                              "
	dim elements
	set elements = getElementsFromQuery(sqlGetData)
	'return
	set getNamedElementsNotOnDiagram = elements
end function

main