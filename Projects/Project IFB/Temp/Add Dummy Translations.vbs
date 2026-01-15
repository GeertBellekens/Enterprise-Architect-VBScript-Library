'[path=\Projects\Project IFB\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Create Dummy Translations
' Author: Geert Bellekens
' Purpose: Create translations for nl and fr for all elements and connectors in the selected package
' Date: 2025-11-26
'


'name of the output tab
const outPutName = "Add Dummy Translations"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp for start
	Repository.WriteOutput outPutName,now() & " Start " & outPutName  , 0
	'actually do the work
	addDummyTranslations
	'set timestamp for finish
	Repository.WriteOutput outPutName,now() & " Finished " & outPutName  , 0
end sub

function addDummyTranslations()
	'get selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	'get all elements
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o where o.Package_ID in (" & packageTreeIDString & ")"
	dim elements
	set elements = getElementsFromQuery(sqlGetData)
	'set dummy translation on elements
	dim element as EA.Element
	for each element in elements
		setDummyTranslation element
	next
	'refresh
	Repository.ReloadPackage package.PackageID
end function

function setDummyTranslation(item)
	dim name
	name = item.Name
	if len(name) = 0 then 
		name = "-anonymous-"
	end if
	Repository.WriteOutput outPutName,now() & " Adding dummy translation for  '" & name & "'", 0
	item.SetTXName "nl", name & " - NL"
	item.SetTXName "fr", name & " - FR"
	item.SetTXNote "nl", name & " - Notes - NL"
	item.SetTXNote "fr", name & " - Notes - FR"
	item.Update
end function

main