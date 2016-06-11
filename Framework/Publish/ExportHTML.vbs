'[path=\Framework\Publish]
'[group=Publish]

option explicit

'
' Script Name: ExportHTML
' Author: Geert Bellekens
' Purpose: Export the model HTML format. This script is suitable to be executed as a scheduled task in order to export the model
'		to HTML and publish it on a webserver or sharepoint site.
' Date: 09/06/2016
'
sub main
	dim repository
	dim projectInterface
	set repository = CreateObject("EA.Repository")
	
	'the path to put the inital export to
	dim exportPath 
	exportPath = "C:\temp\EAExport"
	
	' the path where the exported model should be copied to (sharepoint location, or webserver)
	' in case of a sharepoint location make sure to use the UNC path (\\sharepoint-site\location\) 
	' and make sure that section of sharepoint not version controlled (no check-in/checkout)
	dim publishPath
	publishPath = "C:\temp\copiedFolder"
	
	'the path to the eap file
	dim eapPath
	eapPath = "C:\temp\TMF SQL Server shortcut.EAP"
	
	'open the model
	repository.OpenFile eapPath

	'get project interface
	set projectInterface = repository.GetProjectInterface()
	
	dim packageGUID
	dim rootPackage 
	for each rootPackage in repository.Models
		packageGUID = projectInterface.GUIDtoXML(rootPackage.PackageGUID)
		projectInterface.RunHTMLReport packageGUID, exportPath, ".png", "<default>" , ".html"
		exit for
	next
	'close the model
	repository.CloseFile
	repository.Exit
	
	'copy the export to sharepoint or webserver location
	dim fileSystemObject
	set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
	fileSystemObject.CopyFolder exportPath, publishPath
end sub

main