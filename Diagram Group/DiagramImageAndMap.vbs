'[group=Diagram Group]
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
	set currentDiagram = Repository.GetDiagramByGuid("{D66BC2EC-4D2B-41b2-A29D-4FE3B051EFBD}")
	if not currentDiagram is nothing then
		dim xmlGuid
		dim project as EA.Project
		set project = Repository.GetProjectInterface()
		xmlGUID = project.GUIDtoXML(currentDiagram.DiagramGUID)
		dim gelukt
		gelukt = project.GetDiagramImageAndMap("{D66BC2EC-4D2B-41b2-A29D-4FE3B051EFBD}","c:\temp\")
		session.output "het is gelukt: " & gelukt
	else
		Session.Prompt "This script requires a diagram to be visible", promptOK
	end if

end sub

OnDiagramScript