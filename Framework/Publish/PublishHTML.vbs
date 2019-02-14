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
	'dim repository
	dim projectInterface
	'set repository = CreateObject("EA.Repository")
	
	'the path to put the inital export to
	dim exportPath 
	exportPath = "Y:\Temp\EA\html"
	
	' the path where the exported model should be copied to (sharepoint location, or webserver)
	' in case of a sharepoint location make sure to use the UNC path (\\sharepoint-site\location\) 
	' and make sure that section of sharepoint not version controlled (no check-in/checkout)
	dim publishPath
	publishPath = "http://sharepoint/sites/documentendatabank/EnterpriseArchitecture/SparxExports"
	

	'get project interface
	set projectInterface = repository.GetProjectInterface()
	
	dim packageGUID
	dim rootPackage 
	rootPackage = "{A5CA6155-C51A-4486-99F2-4A6F97D3C8B3}"
	packageGUID = projectInterface.GUIDtoXML(rootPackage)
	projectInterface.RunHTMLReport packageGUID, exportPath, ".png", "<default>" , ".html"
	
	'copy the export to sharepoint or webserver location
	dim fileSystemObject
	set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
	fileSystemObject.CopyFolder exportPath, publishPath
end sub

main