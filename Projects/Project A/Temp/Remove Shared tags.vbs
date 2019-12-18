'[path=\Projects\Project A\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: remove "Shared" tags
' Author: Geert Bellekens
' Purpose: remove all shared tags for the selected package and all packages underneath
' Date: 2019-01-16
'
sub main
	'get selected package
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage
	if selectedPackage is nothing then 
		exit sub
	end if
	'ask user
	dim response
	response = Msgbox("Do you really want to remove all 'Shared' tagged values classes in'" & selectedPackage.Name & "'?" _
	, vbYesNo + vbQuestion, "Remove tagged values")
	If response = vbYes Then
		'actually do the work
		dim sqlDeleteTags
		sqlDeleteTags = "delete tv                                            " & _
						" from t_objectProperties tv                          " & _
						" inner join t_object o on o.Object_ID = tv.Object_ID " & _
						" where tv.Property = 'Shared'                        " & _
						" and o.Package_ID in (" & getPackageTreeIDString(selectedPackage) & ")" 
		Repository.Execute sqlDeleteTags
	end if
end sub

main