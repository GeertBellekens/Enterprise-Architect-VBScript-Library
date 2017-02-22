'[path=\Projects\Project Bellekens\Bellekens Docgen]
'[group=Bellekens DocGen]
option explicit

!INC Bellekens DocGen.UseCaseDocument

' Script Name: Create Use Case Document
' Author: Geert Bellekens
' Purpose: Create the virtual document for a Use Case Document based on the open diagram
' 			Copy this script in a Diagram Group to call it from the diagram directly.
' Date: 11/11/2015
'

sub OnDiagramScript()
	dim documentsPackage as EA.Package
	'select the package to generate the virtual document in
	Msgbox "Please select the package to generate the virtual document in",vbOKOnly+vbQuestion,"Document Package"
	set documentsPackage = selectPackage()
	if not documentsPackage is nothing then
		' Get a reference to the current diagram
		dim currentDiagram as EA.Diagram
		set currentDiagram = Repository.GetCurrentDiagram()
		if not currentDiagram is nothing then
			createUseCaseDocument currentDiagram, documentsPackage.PackageGUID 
			Msgbox "Select the Master Document and press F8 to generate document",vbOKOnly+vbInformation,"Finished!"
		else
			Session.Prompt "This script requires a diagram to be visible", promptOK
		end if
	end if
end sub

OnDiagramScript