'[path=\Projects\Project DF\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC General Scripts.MergeDuplicatesMain

'
' Script Name: Merge Duplicates
' Author: Geert Bellekens
' Purpose: Merge all duplicates within this package branch
' Date: 2025-10-15
'


function Main ()
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	if not selectedPackage is Nothing then
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'inform user
		Repository.WriteOutput outPutName, now() & " Starting " & outPutName &" '" & selectedPackage.Name & "'" , 0
		'do the actual work
		mergeAllDuplicatesInPackage selectedPackage
		'inform user
		Repository.WriteOutput outPutName, now() & " Finsihed " & outPutName &" '" & selectedPackage.Name & "'" , 0
	end if
end function

main

