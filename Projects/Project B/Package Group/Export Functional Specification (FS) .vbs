'[path=\Projects\Project B\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Baloise Scripts.MessageDetailsMain

' Script Name: Export Functional Specification (FS) 
' Author: Geert Bellekens
' Purpose: Export the contents of the messages as a Functional Specification (FS)
' Date: 2023-12-01
'


const outPutName = "Export Functional Specification (FS)"

sub main()
	getMessageDetailsMain functionalDesign
end sub

main
