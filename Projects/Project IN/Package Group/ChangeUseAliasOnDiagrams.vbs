'[path=\Projects\Project IN\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: ChangeUseAliasOnDiagrams
' Author: Geert Bellekens
' Purpose: Change the property Use Alias if Available on the diagrams under this package branche
' Date: 2025-07-05
'
const outPutName = "Change Use Alias"


function main ()
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	if selectedPackage is nothing then 
		exit function
	end if
	if not selectedPackage is Nothing then
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'inform user
		Repository.WriteOutput outPutName, now() & " Starting " & outputName & " for package '" & selectedPackage.Name & "'" , 0
		'do the actual work
		changeUseAlias selectedPackage
		'inform user
		Repository.WriteOutput outPutName, now() & " Starting " & outputName & " for package '" & selectedPackage.Name & "'" , 0
	end if
end function

function changeUseAlias(package)
	dim diagrams
	set diagrams = getAllDiagramsFromPackageBranche(package)
	if diagrams.count = 0 then
		exit function
	end if
	'get user input
	dim response
	response = Msgbox("Set 'Use Alias' to true (Yes) or False (No)?", vbYesNo+vbQuestion, "Use Alias?")
	dim useAliasProperty
	if response = vbYes then
		useAliasProperty = "1"
	else
		useAliasProperty = "0"
	end if
	'loop diagrams
	dim diagram as EA.Diagram
	for each diagram in diagrams
		setUseAlias diagram, useAliasProperty
	next
end function

function setUseAlias(diagram, useAliasProperty)
	dim inverseProperty
	if useAliasProperty = "1" then
		inverseProperty = "0"
	else
		inverseProperty = "1"
	end if
	dim useAliasString
	useAliasString = "UseAlias="
	 if not diagram is nothing then
		dim tempExtendedStyle
		tempExtendedStyle = diagram.ExtendedStyle
		if instr(diagram.ExtendedStyle, useAliasString) > 0 then
			diagram.ExtendedStyle = replace(diagram.ExtendedStyle,useAliasString & inverseProperty,useAliasString & useAliasProperty)
		else
			diagram.ExtendedStyle = diagram.ExtendedStyle & useAliasString & useAliasProperty & ";"
		end if
		if diagram.ExtendedStyle <> tempExtendedStyle then
			diagram.Update
			Repository.WriteOutput outPutName, now() & " Updated Use Alias for diagram '" & diagram.Name & "'" , 0
		end if
	 end if
end function

function getAllDiagramsFromPackageBranche(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData = "select d.Diagram_ID from t_diagram d                   " & vbNewLine & _
				 " where d.Package_ID in (" & packageTreeIDString & ")   "
	dim result
	set result = getDiagramsFromQuery(sqlGetData)
	'return
	set getAllDiagramsFromPackageBranche = result
end function

main