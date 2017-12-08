'[path=\Projects\Project A\Diagram Group]
'[group=Diagram Group]
option explicit

!INC Atrias Scripts.Functional Design Document Main
'
' Script Name: Functional Analysis Document
' Author: Geert Bellekens
' Purpose: Create the virtual document for a functional Analysis document based on the open diagram
' Date: 08/05/2015
'

'
' Diagram Script main function
'
'***update GUID to match the GUID of the FD document packages of this model***
dim FDDocumentsPackageGUID
FDDocumentsPackageGUID = "{8509F5B2-7238-4d13-A9D7-19AB73BDF4EA}"
'***

sub OnDiagramScript()

	' Get a reference to the current diagram
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetCurrentDiagram()
	if not currentDiagram is nothing then
		createFADocument(currentDiagram)
		Msgbox "Finished!"
	else
		Session.Prompt "This script requires a diagram to be visible", promptOK
	end if
end sub

OnDiagramScript

sub test()

	' Get a reference to the current diagram
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetDiagramByGuid("{FE0B0980-389F-4a99-805B-865FEA8CF67A}")
	if not currentDiagram is nothing then
		createFADocument( currentDiagram)
		Msgbox "Finished!"
	else
		Session.Prompt "This script requires a diagram to be visible", promptOK
	end if
end sub

'test