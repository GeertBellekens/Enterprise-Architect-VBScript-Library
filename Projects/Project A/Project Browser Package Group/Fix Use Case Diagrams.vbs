'[path=\Projects\Project A\Project Browser Package Group]
'[group=Project Browser Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Fix Use Case Diagrams
' Author: Geert Bellekens
' Purpose: Checks all use case diagrams under this package and make sure that all the directly used actor are displayed on the diagram, 
' and that the actors that are not linked to any of the use cases on this diagram are removed from the diagram
' Date: 2017-02-24
'
const outputTabName = "Fix Use Case Diagrams"

sub main
	Repository.CreateOutputTab outputTabName
	Repository.ClearOutput outputTabName
	Repository.EnsureOutputVisible outputTabName
	'tell the user we are starting
	Repository.WriteOutput outputTabName, now() & " Starting Fix Use Case Diagrams" ,0
	'get the package id's of the currently selectd package tree
	dim currentPackageTreeIDString
	currentPackageTreeIDString = getCurrentPackageTreeIDString()
	'get all use case diagrams under this package
	dim useCaseDiagrams
	set useCaseDiagrams = getAllUseCaseDiagrams(currentPackageTreeIDString)
	'loop the diagram and fix them
	dim useCaseDiagram as EA.Diagram
	for each useCaseDiagram in useCaseDiagrams
		Repository.WriteOutput outputTabName, now() & " Fixing diagram : " & useCaseDiagram.Name ,0
		fixUseCaseDiagram useCaseDiagram
	next
	'tell the user we are finished
	Repository.WriteOutput outputTabName, now() & " Finished Fix Use Case Diagrams" ,0
end sub

function getAllUseCaseDiagrams(currentPackageTreeIDString)
	dim getUseCaseDiagramsSQL
	getUseCaseDiagramsSQL = "select d.Diagram_ID from t_diagram d         " & _
							" where d.Diagram_Type = 'Use Case'           " & _
							" and d.Package_ID in (" & currentPackageTreeIDString & ")"
								
	set getAllUseCaseDiagrams = getDiagramsFromQuery(getUseCaseDiagramsSQL)
end function

function fixUseCaseDiagram(useCaseDiagram)
	'get the actors that should be on the diagram but are not
	dim primaryActorsToAdd
	set primaryActorsToAdd = getPrimaryActorsToAdd(useCaseDiagram)
	'add the actors to the diagram
	addActorsToDiagram primaryActorsToAdd,useCaseDiagram
	'get the actors that are on the diagram but shouldn't (they don't have a link with anything else on the diagram
	dim actorsToRemove
	set actorsToRemove = getActorsToRemove(useCaseDiagram)
	'remove the unneeded actors from the diagram
	removeActorsFromDiagram actorsToRemove,useCaseDiagram
	'reload the diagram
	Repository.ReloadDiagram useCaseDiagram.DiagramID
	'open the diagram if not already open
	Repository.OpenDiagram useCaseDiagram.DiagramID
end function

function addActorsToDiagram(primaryActorsToAdd,useCaseDiagram)
	'add the actors to the diagram
	dim positionString
	dim x, y, xIncrement, yIncrement
	'set the parameters for adding new elements tot he diagram
	x = 50
	y = 10
	xIncrement = 40
	yIncrement = 20
	'create a new diagramObject for the actor
	dim actor as EA.Element
	for each actor in primaryActorsToAdd
		dim diagramObject as EA.DiagramObject
		positionString =  "l=" & x & ";r=" & x + 45 & ";t=" & y & ";b=" & y +90 & ";"
		'add increments to position
		x = x + xIncrement
		y = y + yIncrement
		'inform the user
		Repository.WriteOutput outputTabName, now() & " Adding actor '" & actor.Name & "' to diagram '" & useCaseDiagram.Name & "'" ,actor.ElementID
		set diagramObject = useCaseDiagram.DiagramObjects.AddNew( positionString, "" )		
		diagramObject.ElementID = actor.ElementID
		'save diagramobject
		diagramObject.Update
	next
end function

function removeActorsFromDiagram(actorsToRemove,useCaseDiagram)
	'loop diagramobjects backwards
	dim i
	dim counter
	counter = 0
	for i = useCaseDiagram.DiagramObjects.Count -1 to 0 step -1
		'stop if we have treated all actors int the list
		if counter = actorsToRemove.Count then
			exit for
		end if
		'loop the actors
		dim actor as EA.Element
		dim diagramObject as EA.DiagramObject
		set diagramObject = useCaseDiagram.DiagramObjects.GetAt(i)
		for each actor in actorsToRemove
			if diagramObject.ElementID = actor.ElementID then
				'found the actor to remove
				'inform the user
				Repository.WriteOutput outputTabName, now() & " Removing actor '" & actor.Name & "' from diagram '" & useCaseDiagram.Name & "'" ,actor.ElementID
				useCaseDiagram.DiagramObjects.DeleteAt i,false
				exit for
			end if
		next
	next
end function


function getPrimaryActorsToAdd(useCaseDiagram)
	dim getActorsSQL
	getActorsSQL = 	"select act.Object_ID from ((t_object act                                              " & _
					" inner join t_connector act_uc on act_uc.Start_Object_ID = act.Object_ID              " & _
					" 								and act_uc.Connector_Type in ('UseCase','Association'))" & _
					" inner join t_object uc on act_uc.End_Object_ID = uc.Object_ID                        " & _
					" 							and uc.Object_Type = 'UseCase' )                           " & _
					" inner join t_diagramobjects do on do.Object_ID = uc.Object_ID                        " & _
					" where act.Object_Type = 'Actor'                                                      " & _
					" and do.Diagram_ID = " & useCaseDiagram.DiagramID & "                                 " & _
					" and not exists (select doa.Instance_ID from t_diagramobjects doa                     " & _
					" 				where doa.Object_ID = act.Object_ID                                    " & _
					" 				and doa.Diagram_ID = " & useCaseDiagram.DiagramID & " )                "
	set getPrimaryActorsToAdd = getElementsFromQuery(getActorsSQL)
end function

function getActorsToRemove(useCaseDiagram)
	dim getActorsSQL
	getActorsSQL = 	"select * from (t_object act                                                    " & _
					" inner join t_diagramobjects act_do on act.Object_ID = act_do.Object_ID)       " & _
					" where                                                                         " & _
					" act.Object_Type = 'Actor'                                                     " & _
					" and not exists                                                                " & _
					" (                                                                             " & _
					" select * from ((t_connector c                                                 " & _
					" inner join t_object o on (o.Object_ID in (c.Start_Object_ID, c.End_Object_ID) " & _
					" 						and o.Object_ID <> act.Object_ID))                      " & _
					" inner join t_diagramobjects do on (do.Object_ID = o.Object_ID                 " & _
					" 								and do.Diagram_ID = act_do.Diagram_ID))         " & _
					" where act.Object_ID in (c.Start_Object_ID, c.End_Object_ID)                   " & _
					" )                                                                             " & _
					" and act_do.Diagram_ID = " &useCaseDiagram.DiagramID & "                       "
	set getActorsToRemove = getElementsFromQuery(getActorsSQL)
end function


main