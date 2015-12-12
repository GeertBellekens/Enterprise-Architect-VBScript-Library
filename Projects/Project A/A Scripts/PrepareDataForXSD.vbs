'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	dim sqlUpdate
	'set all DataTypes to Class
	sqlupdate = "update t_object set object_type = 'Datatype' where object_type = 'Class' and [Package_ID] =" & selectedPackage.PackageID
	Repository.Execute sqlupdate
'	'Set all Enumerations to Class with enumeration stereotype
'	sqlupdate = "update t_object set object_type = 'Class', stereotype = 'enumeration' where object_type = 'Enumeration' and [Package_ID] =" & selectedPackage.PackageID
'	Repository.Execute sqlupdate
	Repository.RefreshModelView selectedPackage.PackageID 
end sub

main