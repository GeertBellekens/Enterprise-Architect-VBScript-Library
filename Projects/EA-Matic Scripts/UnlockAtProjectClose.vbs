'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: UnlockAtProjectClose
' Author: Geert Bellekens
' Purpose: Unlock all locks when closing a project
' Date: 2016-08-05
'
'EA-Matic

function EA_FileClose()
	if Repository.IsSecurityEnabled then
		'get current user id
		dim currentUserID
		currentUserID = Repository.GetCurrentLoginUser(true)
		'figure out how many locks he has
		dim currentUserLocks
		currentUserLocks = getCurrentUserLocks(currentUserID)
		if currentUserLocks > 0 then
			dim response
			response = Msgbox("Unlock all " & currentUserLocks & " locked elements?", vbYesNo+vbQuestion, "Unlock Elements")
			If response = vbYes Then
				dim sqlUnlock
				sqlUnlock = "delete from t_seclocks where UserID = '" & currentUserID & "'"
				Repository.Execute sqlUnlock
			End If
		end if
	end if
end function

function getCurrentUserLocks(currentUserID)
	dim sqlGetLocks 
	sqlGetLocks = "select count(EntityID) AS UserLocks from t_seclocks where UserID = '" & currentUserID & "'"
	dim queryResponse
	queryResponse = Repository.SQLQuery(sqlGetLocks)
    Dim xDoc 
    Set xDoc = CreateObject( "MSXML2.DOMDocument" )
	xDoc.LoadXML(queryResponse)
	dim countNode
	set countNode = xDoc.SelectSingleNode("//UserLocks")
	'return count as integer
	getCurrentUserLocks = CInt(countNode.Text)
end function