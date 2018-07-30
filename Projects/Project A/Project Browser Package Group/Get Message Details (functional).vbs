'[path=\Projects\Project A\Project Browser Package Group]
'[group=Project Browser Package Group]
option explicit

!INC Atrias Scripts.GetMessageDetailsMain
!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' Script Name: Get the message details and export them into an excel file 
' Author: Geert Bellekens
' Purpose: Get Message Details for all messages in this folder and the subfolders and save them to excel
' Date: 2017-03-15
'

'name of the output tab
const outPutName = "Get Message Details (functional)"

sub main
	getMessageDetailsMain false
end sub
main