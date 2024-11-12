'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	Session.Output package.Elements.Count
	dim newElement as EA.Element
	set newElement = package.Elements.AddNew("NewElementName", "Class")
	dim attribute as EA.Attribute
	set attribute = newElement.Attributes.AddNew("attr1", "string")
	attribute.Update
	newElement.Update
	package.Elements.update
	Session.Output package.Elements.Count
'	
'	dim i
'	i = 0
'	dim sqlGetdata
'	sqlGetdata = "select a.ID from t_attribute a                                " & vbNewLine & _
'				" inner join t_object o on o.Object_ID = a.Object_ID           " & vbNewLine & _
'				" inner join t_package p on p.Package_ID = o.Package_ID        " & vbNewLine & _
'				" where p.ea_guid = '" & package.PackageGUID & "'              " & vbNewLine & _
'				" and a.Name like '%object_ID%'                                "
'	dim attributes
'	set attributes = getAttributesFromQuery(sqlGetdata)
'
'	for each attribute in attributes
'		i = i + 1
'		Session.Output "Attribute Name: '" & attribute.Name & "'"
'		attribute.name = attribute.name & "changed"
'		attribute.update
'	next

	'Session.Output "total attributes: " & i
	
end sub

main