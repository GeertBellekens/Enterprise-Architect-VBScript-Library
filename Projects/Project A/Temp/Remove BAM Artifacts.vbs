'[path=\Projects\Project A\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Remove BAM artifacts
' Author: Geert Bellekens
' Purpose: Remove all BAM artifacts from the selected package branch
' Date: 2019-08-08
'

'name of the output tab
const outPutName = "Remove BAM artifacts"

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
	userinput = MsgBox( "Remove all BAM artifacts from package '"& selectedPackage.Name &"'?", vbYesNo + vbQuestion, "Remove all BAM artifacts?")
	'save the schema content
	if userinput = vbYes then
		'report progress
		Repository.WriteOutput outPutName, now() & " Starting remove BAM artifacts for package '"& selectedPackage.Name &"'", 0
		'actually remove the BAM artifacts
		removeBAMartifacts selectedPackage
		'report progress
		Repository.WriteOutput outPutName, now() & " Finished remove BAM artifacts for package '"& selectedPackage.Name &"'", 0
	end if 
end sub

function removeBAMartifacts(package)
	'get the package tree id's
	dim packageTreeIDs
	packageTreeIDs = getPackageTreeIDString(package)
	dim sqlGetBamArtifacts
	sqlGetBamArtifacts = "select o.Object_ID from t_object o            " & _
						 " where o.Stereotype = 'BAM_Specification'     " & _
						 " and o.Package_ID in (" & packageTreeIDs & ") "
	dim bamArtifacts
	set bamArtifacts = getElementsFromQuery(sqlGetBamArtifacts)
	'loop bamArtifacts and delete each of them
	dim bamArtifact as EA.Element
	for each bamArtifact in bamArtifacts
		removeBamArtifact bamArtifact
	next
end function

function removeBamArtifact(bamArtifact)
	'get the owner of the BAM artifact
	dim owner as EA.Element
	set owner = Repository.GetElementByID(bamArtifact.ParentID)
	dim i
	i = 0
	dim candidate as EA.Element
	for each candidate in owner.Elements
		if candidate.ElementID = bamArtifact.ElementID then
			'report progress
			Repository.WriteOutput outPutName, now() & " Deleting BAM artifact '"& bamArtifact.Name &"'", bamArtifact.ElementID
			owner.Elements.DeleteAt i, false
			exit for
		end if
		'not found, up the counter
		i = i + 1		
	next
end function

main