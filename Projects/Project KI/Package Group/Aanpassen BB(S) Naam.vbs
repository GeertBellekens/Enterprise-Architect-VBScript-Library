'[path=\Projects\Project KI\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Aanpassen BB(S) Naam
' Author: Geert Bellekens
' Purpose: Kopieer de waarden van de tagged values BBSNaam en BBNaam naar de onderliggende packages (zonder stereotype)
' Date: 2025-12-03
'
const outPutName = "Aanpassen BB(S) Naam"
const tagNames = "BBNaam,BBSNaam"


function main ()
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	if not selectedPackage is Nothing then
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'inform user
		Repository.WriteOutput outPutName, now() & " Starting " & outPutName & "for package '" & selectedPackage.Name & "'" , 0
		'do the actual work
		updatePackageTags selectedPackage
		'inform user
		Repository.WriteOutput outPutName, now() & " Finished " & outPutName & "for package '" & selectedPackage.Name & "'" , 0
	end if
end function

function updatePackageTags(package)
	'get starting tag values to copy down
	dim tagDictionary
	set tagDictionary = CreateObject("Scripting.Dictionary")
	dim tagNamesList
	tagNamesList = Split(tagNames, ",")
	dim tagName
	for each tagName in tagNamesList
		dim tagValue
		tagValue = getTaggedValueValue(package.Element, tagName)
		tagDictionary.Add tagName, tagValue
	next
	'copy the tag values down
	copyTagValuesDown package, tagDictionary
end function

function copyTagValuesDown(package, tagDictionary)
	'only update tags if stereotype is empty
	if len(package.StereotypeEx) = 0 then
		Repository.WriteOutput outPutName, now() & " Updating tags for package '" & package.Name & "'" , 0
		dim tagName
		for each tagName in tagDictionary.Keys
			dim tagValue
			tagValue = tagDictionary(tagName)
			dim tag as EA.TaggedValue
			set tag = getExistingOrNewTaggedValue(package.Element, tagname)
			tag.Value = tagValue
			tag.Update
		next
	end if
	'loop subPackages
	dim subPackage as EA.Package
	for each subPackage in package.Packages
		copyTagValuesDown subPackage, tagDictionary
	next
end function

main