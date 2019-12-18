'[path=\Projects\Project A\Model Management]
'[group=Model Management]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Clear All Notes in Selected Package
' Author: Geert Bellekens
' Purpose: Clears the notes for all elements, attributes and relations for the selected package and everything underneath
' Date: 2019-04-09
'
const outPutName = "Clear All Notes"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get selected package
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage
	if not selectedPackage is nothing then
		'ask for confirmation
		dim userIsSure
		userIsSure = Msgbox("Are you sure you want clear all notes in the package '" &selectedPackage.Name & "' and all its subPackages? '", vbYesNo+vbExclamation, "Clear Notes in Package " &selectedPackage.Name & "?")
		if userIsSure = vbYes then
			Repository.WriteOutput outPutName, now() & " Starting clear notes for package '"& selectedPackage.Name &"'", 0
			'delete the package using it's parent
			clearNotes selectedPackage
			'let user know
			Repository.WriteOutput outPutName, now() & " Finished clear notes for package '"& selectedPackage.Name &"'", 0
		end if
	end if
end sub

function clearNotes(selectedPackage)
	dim packageTreeIDs
	packageTreeIDs = getPackageTreeIDString(selectedPackage)
	'inform user
	Repository.WriteOutput outPutName, now() & " Clearing Element notes", 0
	dim sqlClearElementNotes
	sqlClearElementNotes = "update o set o.Note = null                         " & _
						   " from t_object o                                   " & _
						   " where o.Package_ID in (" & packageTreeIDs & ")    "
    Repository.Execute sqlClearElementNotes
	'inform user
	Repository.WriteOutput outPutName, now() & " Clearing Attribute notes", 0
	dim sqlClearAttributeNotes
	sqlClearAttributeNotes = "update a set a.Notes = null                          " & _
							" from t_attribute a                                  " & _
							" inner join t_object o on o.Object_ID = a.Object_ID  " & _
							" where o.Package_ID in (" & packageTreeIDs & ")      "
	Repository.Execute sqlClearAttributeNotes
	'inform user
	Repository.WriteOutput outPutName, now() & " Clearing Relationship notes", 0
	dim sqlClearRelationNotes
	sqlClearRelationNotes = "update c set c.Notes = null                               " & _
							" from t_connector c                                       " & _
							" inner join t_object o on o.Object_ID = c.Start_Object_ID " & _
							" where o.Package_ID in (" & packageTreeIDs & ")           "
	Repository.Execute sqlClearRelationNotes
end function

main