'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: some text here
' Author: 
' Purpose: 
' Date: 
'
sub main
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage
	dim diagram as EA.Diagram
	set diagram = package.Diagrams.AddNew("newDiagramName", "Logical")
	diagram.Update
	dim element as EA.Element
	set element = package.Elements.AddNew("className","Class")
	element.Update
	dim diagramObject as EA.DiagramObject
	set diagramObject = diagram.DiagramObjects.AddNew("","")
	diagramObject.ElementID = element.ElementID
	diagramObject.Update
	Repository.ReloadDiagram(diagram.DiagramID)
	
end sub

main