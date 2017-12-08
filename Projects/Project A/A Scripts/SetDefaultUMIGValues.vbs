'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: SetDefaultUMIGValues
' Author: Geert Bellekens
' Purpose: Sets the default value for the UMIG tagged values for the selected package
' Date: 2017-04-19
'
'name of the output tab
const outPutName = "Set Default UMIG Values"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName

	'ask the user to select a package
	msgbox "Please select the package containing the FIS's"
	dim userSelectedPackage as EA.Package
	set userSelectedPackage = selectPackage()
	if not userSelectedPackage is nothing then
		'set timestamp for start
		Repository.WriteOutput outPutName,now() & " Starting Set default Value for UMIG tagged Value"  , 0
		dim defaultValue
		defaultValue = InputBox("Please enter the default Value for the UMIG tagged value","UMIG Default Value")
		if len(defaultValue) > 0 then
			'get packageTreeIDstring
			dim packageTreeID
			packageTreeID = getPackageTreeIDString(userSelectedPackage)
			dim updateUmigQuery
			 updateUmigQuery = "update tv set tv.Value = '" & defaultValue & "'     " & _
								" from t_objectproperties tv                         " & _
								" inner join t_object o on tv.Object_ID = o.Object_ID" & _
								" where  tv.Property = 'Atrias::UMIG'                " & _
								" and o.Package_ID in (" & packageTreeID & ") 		"
			Repository.Execute updateUmigQuery
		end if
		'set timestamp for end
		Repository.WriteOutput outPutName,now() & " Finished Set default Value for UMIG tagged Value"  , 0
	end if

end sub


main