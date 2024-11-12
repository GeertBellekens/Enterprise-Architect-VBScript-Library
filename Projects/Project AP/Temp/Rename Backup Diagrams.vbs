'[path=\Projects\Project AP\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Rename Backup Diagrams
' Author: Geert Bellekens
' Purpose: Add the prefix "zzBackup_" to all of the diagram in a package ending with "_backup"
' Date: 2022-05-27
'
const outPutName = "Rename backup diagrams "

function Main ()
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	if selectedPackage is nothing then
		exit function
	end if
	'inform user
	Repository.WriteOutput outPutName, now() & " Starting Rename backup diagrams for package '" & selectedPackage.name & "'", 0
	'do the actual work
	renameBackupDiagrams selectedPackage
	'inform user
	Repository.WriteOutput outPutName, now() & " Finished Rename backup diagrams for package '" & selectedPackage.name & "'", 0
		
end function

function renameBackupDiagrams(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData = "select d.Diagram_ID from t_diagram d                    " & vbNewLine & _
				" inner join t_package p on p.Package_ID = d.Package_ID   " & vbNewLine & _
				" where p.Name like '%_backup'                            " & vbNewLine & _
				" and p.Package_ID in (" & packageTreeIDString & ")       " & vbNewLine & _
				" and not d.name like 'zzBackup_%'"
	dim results
	set results = getDiagramsFromQuery(sqlGetData)
	Repository.WriteOutput outPutName, now() & " Found " & results.Count & " diagrams to rename", 0
	if results.Count = 0 then
		'no diagrams found
		exit function
	end if
	'make sure the user is sure:
	dim userIsSure
	userIsSure = Msgbox("Rename " & results.Count & " diagrams in package '" & package.Name & "'?", vbYesNo+vbQuestion, "Rename Backup Diagrams?")
	if userIsSure = vbYes then
		dim diagram as EA.Diagram
		dim i
		i = 0
		for each diagram in results
			i = i + 1
			'inform user
			Repository.WriteOutput outPutName, now() & " Renaming diagram " & i & "of " & results.Count & ": '" & diagram.Name & "'", 0
			diagram.Name = "zzBackup_" & diagram.Name
			diagram.Update
		next
	end if
end function

main