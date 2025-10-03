'[path=\Projects\Project I\Project Browser Group]
'[group=Project Browser Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' Script Name: Translate Elements
' Author: Geert Bellekens
' Purpose: Translate the selected elements in the project browser
' Date: 2025-08-27

const outPutName = "Translate Elements"


sub main
	translateProjectBrowser()
end sub

main