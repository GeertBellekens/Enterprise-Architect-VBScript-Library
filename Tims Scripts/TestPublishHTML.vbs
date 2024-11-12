'[group=Tims Scripts]
option explicit

'!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
 ' TODO: Enter script code here!
 dim projectInterface
 set projectInterface = repository.GetProjectInterface()

 dim exportPathS
 exportPathS = "H:\_Projecten\EA\Sparx\htmlexports\services"
 dim packageGUIDS
 dim rootPackageS
 rootPackageS = "{74CAF8D0-A65A-4e59-B810-DB57596A28F5}"
 packageGUIDS = projectInterface.GUIDtoXML(rootPackageS)
 projectInterface.RunHTMLReport packageGUIDS, exportPathS, ".png", "<default>" , ".html"

end sub

main