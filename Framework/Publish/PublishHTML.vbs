'[path=\Framework\Publish]
'[group=Publish]

option explicit
!INC Local Scripts.EAConstants-VBScript
'
' Script Name: ExportHTML
' Original Author: Geert Bellekens
' Purpose: Export the model HTML format. This script is suitable to be executed as a scheduled task in order to export the model
'  to HTML and publish it on a webserver or sharepoint site.
' Date: 09/06/2016
'
sub main
 'the paths to put the inital export to
 dim exportPathA 
 exportPathA = "H:\_Projecten\EA\Sparx\htmlexports\applicaties"
 dim exportPathI 
 exportPathI = "H:\_Projecten\EA\Sparx\htmlexports\integraties"
 dim exportPathS 
 exportPathS = "H:\_Projecten\EA\Sparx\htmlexports\services"
 
 ' the path where the exported model should be copied to (sharepoint location, or webserver)
 ' in case of a sharepoint location make sure to use the UNC path (\\sharepoint-site\location\) 
 ' and make sure that section of sharepoint not version controlled (no check-in/checkout)
 dim publishPathA
 publishPathA = "\\sharepoint\DavWWWRoot\sites\documentendatabank\EnterpriseArchitecture\SparxExports\applicaties"
 dim publishPathI
 publishPathI = "\\sharepoint\DavWWWRoot\sites\documentendatabank\EnterpriseArchitecture\SparxExports\integraties"
 dim publishPathS
 publishPathS = "\\sharepoint\DavWWWRoot\sites\documentendatabank\EnterpriseArchitecture\SparxExports\services"
 
 'copy the export to sharepoint or webserver location
 dim fileSystemObject
 set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
 fileSystemObject.CopyFolder exportPathA, publishPathA
 'fileSystemObject.CopyFolder exportPathI, publishPathI
 'fileSystemObject.CopyFolder exportPathS, publishPathS
end sub

main