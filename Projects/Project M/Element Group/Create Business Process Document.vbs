'[path=\Projects\Project M\Element Group]
'[group=Element Group]
'[group_type=CONTEXTELEMENT]

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Metallo Modelling Standards.Business Process Document Main
'
' Script Name: Create Business Process Document
' Author: Geert Bellekens
' Purpose: Call the main function to create the virtual document for the Business Process Document based on the selected Archimate Process 
' Date: 2017-02-16
'

'update this value to correspond to the package where the virtual documents should be created
const businessProcessDocumentsPackageGUID = "{6C6AFA41-B06C-4d65-B11C-1816EF5811CD}"

sub main
	'get the selected element
	dim rootBusinessProcess
	set rootBusinessProcess = Repository.GetContextObject()
	'call the main function
	createNewBusinessProcessDocument businessProcessDocumentsPackageGUID, rootBusinessProcess
end sub

main