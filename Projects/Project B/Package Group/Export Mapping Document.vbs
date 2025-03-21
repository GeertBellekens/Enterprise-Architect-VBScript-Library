'[path=\Projects\Project B\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Baloise Scripts.MessageDetailsMain

' Script Name: Export Mapping Document
' Author: Geert Bellekens
' Purpose: Export the contents of the messages as a Mapping Document
' Date: 2023-12-01
'


const outPutName = "Export Mapping Document"

sub main()
	getMessageDetailsMain mappingDocument
end sub

main