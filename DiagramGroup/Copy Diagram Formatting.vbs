'[group=DiagramGroup]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' This code has been included from the default Diagram Script template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.
'
' Script Name:
' Author:
' Purpose:
' Date:
'

'
' Diagram Script main function
'
sub OnDiagramScript()

	' Get a reference to the current diagram
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetCurrentDiagram()

	if not currentDiagram is nothing then
		' Get a reference to any selected connector/objects
		dim selectedConnector as EA.Connector
		dim selectedObjects as EA.Collection
		set selectedConnector = currentDiagram.SelectedConnector
		set selectedObjects = currentDiagram.SelectedObjects

		if not selectedConnector is nothing then
			' A connector is selected
		elseif selectedObjects.Count > 0 then
			' One or more diagram objects are selected
		else
			' Nothing is selected
		end if
	else
		Session.Prompt "This script requires a diagram to be visible", promptOK
	end if

end sub

OnDiagramScript
