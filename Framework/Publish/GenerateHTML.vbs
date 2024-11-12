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
 dim projectInterface
 
 'the paths to put the inital export to
 dim exportPathA 
 exportPathA = "H:\_Projecten\EA\Sparx\htmlexports\applicaties"
 dim exportPathI 
 exportPathI = "H:\_Projecten\EA\Sparx\htmlexports\integraties"
 dim exportPathS 
 exportPathS = "H:\_Projecten\EA\Sparx\htmlexports\services"
 
 'get project interface
 set projectInterface = repository.GetProjectInterface()
 
 dim packageGUIDA
 dim rootPackageA
 rootPackageA = "{89CC255F-72C3-486f-90C4-20D45E5AD61E}"
 dim packageGUIDI
 dim rootPackageI 
 rootPackageI = "{999E603A-E75B-4f07-A013-87D44B5BBC41}"
 dim packageGUIDS
 dim rootPackageS 
 rootPackageS = "{74CAF8D0-A65A-4e59-B810-DB57596A28F5}"

 packageGUIDA = projectInterface.GUIDtoXML(rootPackageA)
 projectInterface.RunHTMLReport packageGUIDA, exportPathA, ".png", "DeLijn" , ".html"
 packageGUIDI = projectInterface.GUIDtoXML(rootPackageI)
 projectInterface.RunHTMLReport packageGUIDI, exportPathI, ".png", "DeLijn" , ".html"
 packageGUIDS = projectInterface.GUIDtoXML(rootPackageS)
 projectInterface.RunHTMLReport packageGUIDS, exportPathS, ".png", "DeLijn" , ".html"
end sub

main