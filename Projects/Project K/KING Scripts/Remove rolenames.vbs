'[path=\Projects\Project K\KING Scripts]
'[group=KING Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	dim sqlUpdate
	sqlUpdate = "update t_connector set SourceRole = null, DestRole = null " &_
	" where ([SourceRole] like '*' or [DestRole] like '*' )" &_
	" and [Start_Object_ID] <> [End_Object_ID] "
	Repository.Execute sqlUpdate
end sub

main