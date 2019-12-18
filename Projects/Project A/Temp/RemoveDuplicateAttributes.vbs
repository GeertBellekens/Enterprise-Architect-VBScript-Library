'[path=\Projects\Project A\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: RemoveDuplicateAttributes
' Author: Geert Bellekens
' Purpose: Remove the duplicate attributes in the selected package
' Date: 2019-01-03
'
sub main
	dim selectedPackage as EA.Package
	dim packageTreeIDs 
	packageTreeIDs  = getCurrentPackageTreeIDString()
	dim sqlGetAttributes
	sqlGetAttributes = "SELECT a.ID                                   " & _
						" from t_attribute a                                 " & _
						" inner join t_object o on o.Object_ID = a.Object_ID " & _
						" where exists                                       " & _
						" (select a2.ID from t_attribute a2                  " & _
						" where a2.Object_ID = a.Object_ID                   " & _
						" and a2.name = a.name                               " & _
						" and a2.ID < a.id)                                  " & _
						" and o.Package_ID in (" & packageTreeIDs & ")       " & _
						" ORDER BY a.Object_ID, a.name                       "
	dim attributes
	set attributes = getattributesFromQuery(sqlGetAttributes)
	dim attribute as EA.Attribute
	for each attribute in attributes
		dim owner as EA.Element
		set owner = Repository.GetElementByID(attribute.ParentID)
		dim i
		for i = owner.Attributes.Count -1 to  0 step -1
			if owner.attributes(i).AttributeID = attribute.AttributeID then
				Session.Output "Deleting attribute " & attribute.Name & " with GUID " & attribute.AttributeGUID
				owner.Attributes.DeleteAt i,false
			end if
		next
	next
end sub

main