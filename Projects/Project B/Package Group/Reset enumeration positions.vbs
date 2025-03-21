'[path=\Projects\Project B\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' Script Name: Reset enumeration positions
' Author: Geert Bellekens
' Purpose: Sets the positions of the enumeration values in teh selected package to 0, in order to be processed in alphabetical order
' Date: 2019-02-05
'


const outPutName = "Reset Enumeration positions"
sub main()
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	
	'get the selected element
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetContextObject
	if selectedPackage.ObjectType = otPackage then
		dim response
		response = Msgbox("Do you want reset positions for all enumerations in package '" & selectedPackage.Name  & "'?", vbYesNoCancel+vbQuestion, "Reset enumeration positions")
		if response = VBYes then		
			dim sqlUpdatePos
			sqlUpdatePos = "update a set a.Pos = 0                                             " & _
							" from t_attribute a                                                 " & _
							" inner join t_object o on o.Object_ID = a.Object_ID                 " & _
							" 						and (o.Object_Type = 'Enumeration'           " & _
							" 						or isnull(o.stereotype, '') = 'Enumeration') " & _
							" where a.Pos <> 0                                                   " & _
							" and o.Package_ID in (" & getPackageTreeIDString(selectedPackage) & ")"
			'execute sql update
			Repository.Execute sqlUpdatePos
			'reload package to show the new order
			Repository.ReloadPackage(selectedPackage.PackageID)
			'tell the user we are starting
			Repository.WriteOutput outPutName, now() & " Finished resetting enumeration positions for package '" & selectedPackage.Name & "'", selectedPackage.Element.ElementID
		end if 
	end if
end sub

main















