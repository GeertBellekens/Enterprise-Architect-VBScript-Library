'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Check CheckedOut Packages at Project Close
' Author: Geert Bellekens
' Purpose: Check the number of checked out packages when closing the project to remind the user to check them in.
' Date: 2017-10-20
'
'EA-Matic

function EA_FileClose()
	dim userName
	userName = CreateObject("WScript.Network").UserName
	'figure out if there are any packages still checked out
	
	'figure out how many locks he has
	dim checkedOutPackageCount
	checkedOutPackageCount = getNumberOfCheckedOutPackge(userName)
	if checkedOutPackageCount > 0 then
		Msgbox "You have " & checkedOutPackageCount & " packages still checked out." & vbNewLine & _
				"Please check in your changes as soon as possible!", vbOk + vbExclamation, "Checked out packages"
	end if
end function

function getNumberOfCheckedOutPackge(userName)
	dim getGetPackageCount
	getGetPackageCount = "select count(p.Package_ID) as packageCount from t_package p where p.PackageFlags like '%CheckedOutTo=" & userName & ";%' "
	dim queryResponse
	queryResponse = Repository.SQLQuery(getGetPackageCount)
    Dim xDoc 
    Set xDoc = CreateObject( "MSXML2.DOMDocument" )
	xDoc.LoadXML(queryResponse)
	dim countNode
	set countNode = xDoc.SelectSingleNode("//packageCount")
	'return count as integer
	getNumberOfCheckedOutPackge = CInt(countNode.Text)
end function
