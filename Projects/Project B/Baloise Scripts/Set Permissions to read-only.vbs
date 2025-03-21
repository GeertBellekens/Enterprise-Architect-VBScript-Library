'[path=\Projects\Project B\Baloise Scripts]
'[group=Baloise Scripts]

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include


'
' Script Name: Set Permissions to Read-Only
' Author: Geert Bellekens
' Purpose: Removes all permissions of all groups except the Administrators group
' Date: 2025-03-10
'
const outPutName = "Set Permissions to Read-only"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'inform user
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName , 0
	'do the actual works
	setPermissionToReadOnly
	'inform user
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName , 0
end sub

function setPermissionToReadOnly()
		'check if the user is part of the administrators
	dim userID
	userID = Repository.GetCurrentLoginUser(true)
	dim sqlGetData
	sqlGetData = "select ug.UserID from t_secusergroup ug                     " & vbNewLine & _
				" inner join t_secgroup g on g.GroupID = ug.GroupID          " & vbNewLine & _
				" 				and g.GroupName  = 'Administrators'          " & vbNewLine & _
				" and ug.UserID = '" & userID & "'   "
	dim results
	set results = getArrayListFromQuery(sqlGetData)
	if results.Count = 0 then
		Msgbox "This script can only be executed by members of the Administrators group", vbCritical, "Script execution not allowed!"
		exit function
	end if
	'user confirmation
	'ask the user if he is sure
	dim userIsSure
	userIsSure = Msgbox("Do you really want remove all permissions for all groups except the Administrators?", vbYesNo+vbQuestion, "Remove all permissions?")
	if userIsSure = vbYes then
		dim sqlUpdate
		sqlUpdate = "delete p                                            " & vbNewLine & _
				" from t_secgrouppermission p                        " & vbNewLine & _
				" inner join t_secgroup g on g.GroupID = p.GroupID   " & vbNewLine & _
				" where g.GroupName <> 'Administrators'              "
		Repository.Execute sqlUpdate
	end if
end function

main