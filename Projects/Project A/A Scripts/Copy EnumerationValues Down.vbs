'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Copy Enumeration Values Down
' Author: Geert Bellekens
' Purpose: Copy the values of the inherited enumerations down to the sub-enumerations
' Date: 2017-04-19
'
'name of the output tab
const outPutName = "Copy Enumeration values Down"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp for start
	Repository.WriteOutput outPutName,now() & " Starting Copy enumeration Values Down"  , 0
	'ask the user to select a package
	msgbox "Please select the package containing the enumerations"
	dim userSelectedPackage as EA.Package
	set userSelectedPackage = selectPackage()
	if not userSelectedPackage is nothing then
		if isRequireUserLockEnabled() then
			if not userSelectedPackage.ApplyUserLock() then
				msgbox "Please apply user lock to the selected package",vbOKOnly+vbExclamation,"Selected Package not locked!"
				exit sub
			end if
			'copy enumeration values down for enumerations in the selected package
			copyEnumValuesDown userSelectedPackage
		end if
	end if
	'set timestamp for end
	Repository.WriteOutput outPutName,now() & " Finished Copy enumeration Values Down"  , 0
end sub

function copyEnumValuesDown(userSelectedPackage)
	'get the idstring for the package
	dim packageIDString
	packageIDString = getPackageTreeIDString(userSelectedPackage)
	dim getEnumerationsSQL
	getEnumerationsSQL = "select o.Object_ID from t_object o   " & _
						" where                                " & _
						" (o.Object_Type = 'Enumeration'       " & _
						" or                                   " & _
						" (o.Object_Type = 'Class'             " & _
						" and o.Stereotype = 'Enumeration' ))  " & _
						" and o.Package_ID in (" & packageIDString & ")"
	dim enumerations
	set enumerations = getElementsFromQuery(getEnumerationsSQL)
	'loop the enumerations and copy the values of their parents down
	dim enumeration as EA.Element
	for each enumeration in enumerations
		Repository.WriteOutput outPutName,now() & " Processing enumeration: " & enumeration.Name , enumeration.ElementID
		copyValuesDown enumeration
	next
end function

function copyValuesDown(enumeration)
	dim parentEnum as EA.Element
	dim enumvalues
	set enumvalues = CreateObject("Scripting.Dictionary")
	'add all current values tot dictionary
	dim currentValue as EA.Attribute
	for each currentValue in enumeration.Attributes
		if not enumvalues.Exists(currentValue.Name) then
			enumvalues.Add currentValue.Name, currentValue
		end if
	next
	for each parentEnum in enumeration.BaseClasses
		copyValuesFromParent enumeration, parentEnum, enumvalues
	next
end function

function copyValuesFromParent(enumeration, parentEnum, enumvalues)
	dim currentValue as EA.Attribute
	dim parentValue as EA.Attribute
	for each parentValue in parentEnum.Attributes
		if not enumvalues.Exists(parentValue.Name) then
			'create the new value
			dim newValue as EA.Attribute
			set newValue = enumeration.Attributes.AddNew(parentValue.Name,"")
			newValue.notes = parentValue.Notes
			newValue.Update
			'add it to the dictionary
			enumvalues.Add newValue.Name, newValue
		end if
		'then go one level up
		dim grandParentEnum as EA.Element
		for each grandParentEnum in parentEnum.BaseClasses
			copyValuesFromParent enumeration, grandParentEnum
		next
	next
end function

main