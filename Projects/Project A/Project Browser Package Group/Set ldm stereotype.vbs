option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	session.output package.Name
	dim classElement as EA.Element
	if package.Name = "Logical Data Model Entities" then
		for each classElement in package.Elements
			if classElement.Type = "Class" and classElement.Stereotype = vbNullString then
				classElement.Stereotype = "ldm"
				classElement.Update
			end if
		next
	end if
	Repository.RefreshModelView package.PackageID
	msgbox "finished!"
end sub

main