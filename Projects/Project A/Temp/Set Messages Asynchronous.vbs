'[path=\Projects\Project A\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Set Messages Asynchronous
' Author: Geert Bellekens
' Purpose: Sets all messages in the sequence diagrams under the selected package to Asynchronous
' Date: 2018-01-09
'

const outPutName = "Set Messages Asynchronous"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get selected package
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage
	if not selectedPackage is nothing then 
		dim response
		response = Msgbox("Set all messages to Asynchronous for package '" & selectedPackage.Name & "'?" , vbYesNo + vbQuestion, "Set Messages Asynchronous")
		If response = vbYes Then
			'set timestamp
			Repository.WriteOutput outPutName, now() & " Start setting messages asynchronous for '" & selectedPackage.Name &  "'", 0
			'actually do the work
			setMessagesAsynchronous selectedPackage
			'set timestamp
			Repository.WriteOutput outPutName, now() & " Finished setting messages asynchronous for '" & selectedPackage.Name &  "'", 0
		end if
	end if
end sub

function setMessagesAsynchronous(package)
	'test if package is not readonly
	on error resume next
	package.update 'this will return an error if the package is readonly
	if Err.number <> 0 then
		Err.clear
		on error goto 0
		Repository.WriteOutput outPutName, now() & " Skipped package '" & package.Name & "' because it is read-only", 0
		
	else 'no error, package not readonly
		on error goto 0
		dim sqlUpdateMessages
		sqlUpdateMessages = "update c set PDATA1 = 'Asynchronous'                        " & _
							" from (t_connector c                                        " & _
							" inner join t_diagram d on d.Diagram_ID = c.DiagramID)      " & _
							" where d.Package_ID = " & package.PackageID 
		Repository.Execute sqlUpdateMessages
		Repository.WriteOutput outPutName, now() & " Processed package '" & package.Name & "'", 0
	end if
	'process subpackages
	dim subPackage
	for each subPackage in package.Packages
		setMessagesAsynchronous(subPackage)
	next
end function

main