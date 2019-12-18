'[path=\Projects\Project A\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Remove External Refernces
' Author: Geert Bellekens
' Purpose: Remove all External reference boundaries from the selected package branch
' Date: 2019-08-27
'

'name of the output tab
const outPutName = "Remove External references"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get selected package
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage
	if selectedPackage is nothing then
		msgbox "Please select a package in the project browser before running this script",vbOKOnly+vbExclamation,"No package selected!"
		exit sub
	end if
	'get confirmation
	dim userInput
	userinput = MsgBox( "Remove all External References from package '"& selectedPackage.Name &"'?", vbYesNo + vbQuestion, "Remove all External References?")
	'save the schema content
	if userinput = vbYes then
		'report progress
		Repository.WriteOutput outPutName, now() & " Starting remove External References for package '"& selectedPackage.Name &"'", 0
		'actually remove the external references
		removeExternalReferences selectedPackage
		'report progress
		Repository.WriteOutput outPutName, now() & " Finished remove External References for package '"& selectedPackage.Name &"'", 0
	end if 
end sub

function removeExternalReferences(package)
	Repository.WriteOutput outPutName, now() & " Processing Package '"& package.Name &"'", package.Element.ElementID
	dim sqlGetexternalReferences
	sqlGetexternalReferences = "select o.Object_ID from t_object o             " & _
						 " where o.Object_Type = 'Boundary' and o.NType = 1001 " & _
						 " and o.Package_ID = " & package.PackageID & "        "
	dim externalReferenceIDs
	set externalReferenceIDs = getVerticalArrayListFromQuery(sqlGetExternalReferences)
	if externalReferenceIDs.Count > 0 then
		set externalReferenceIDs = externalReferenceIDs(0) 'get the first array
		'create dictionary
		dim externalReferenceDict
		set externalReferenceDict = CreateObject("Scripting.Dictionary")
		dim externalReference
		for each externalReference in externalReferenceIDs
			'add to dictionary
			'Repository.WriteOutput outPutName, now() & " Adding ext. ref. ID:  '"& externalReference &"'", 0
			externalReferenceDict.Add externalReference, externalReference
		next
		'loop external Referenes and delete each of them
		dim i
		i = 0
		dim candidate as EA.Element
		for each candidate in package.Elements
			'check if id is in Dictionary
			if externalReferenceDict.Exists(cstr(candidate.ElementID)) then
				'report progress
				Repository.WriteOutput outPutName, now() & " Deleting External Reference '"& candidate.Name &"'", 0
				package.Elements.DeleteAt i, false
			end if
			'not found, up the counter
			i = i + 1		
		next
	end if
	'loop subPackages
	dim subPackage as EA.Package
	for each subPackage in package.Packages
		removeExternalReferences subPackage
	next
end function

main