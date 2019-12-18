'[path=\Projects\Project A\Schema Scripts]
'[group=Schema Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' Script Name: Set XSD position tags
' Author: Geert Bellekens
' Purpose: Set the XSD position tags for all XSDElement attributes and enumeration values in the selected package
' Date: 2019-09-10
'

'name of the output tab
const outPutName = "Set XSD position tags"

sub main

	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	
	if not selectedPackage is nothing then
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'speed up
		Repository.EnableUIUpdates = false
		'set timestamp
		Repository.WriteOutput outPutName, now() & " Starting set XSD position tags for '"& selectedPackage.Name &"'", 0
		'set the actual position tags
		setXSDPositionTags selectedPackage
		'speed down
		Repository.EnableUIUpdates = true
		'refresh
		Repository.RefreshModelView 0
		'set timestamp
		Repository.WriteOutput outPutName, now() & " Finished set XSD position tags for '"& selectedPackage.Name &"'", 0
	end if
end sub

function setXSDPositionTags(package)
	'loop elements
	dim element as EA.Element
	for each element in package.elements
		'inform user
		Repository.WriteOutput outPutName, now() & " Processing element '"& element.Name &"'", element.ElementID
		'correct positions
		correctAttributePositions element
		'add missing position tags
		addMissingPositionTags element
		'fix wrong position tags
		correctPositionTags element 
	next
	'process sub-packages
	dim subPackage as EA.Package
	for each subPackage in package.Packages
		setXSDPositionTags subPackage
	next
end function

function correctPositionTags(element)
	'update the position tags with values different from the pos
	dim sqlGetAttributes
	sqlGetAttributes = "select a.ID from t_attribute a                            " & vbNewLine & _
						" inner join t_attributetag tv on tv.ElementID = a.ID     " & vbNewLine & _
						" 							and tv.Property = 'position'  " & vbNewLine & _
						" inner join t_object o on o.Object_ID = a.Object_ID      " & vbNewLine & _
						" where a.Object_ID = " & element.ElementID & "           " & vbNewLine & _
						" and (a.Stereotype = 'XSDelement'                        " & vbNewLine & _
						" or  o.Object_Type = 'Enumeration')                      " & vbNewLine & _
						" and isnull(tv.VALUE, 'X') <> CAST(a.pos + 1 as varchar(10)) "
	dim attributes
	set attributes = getAttributesFromQuery(sqlGetAttributes)
	dim attribute as EA.Attribute
	for each attribute in attributes
		dim tv as EA.TaggedValue
		set tv = getExistingOrNewTaggedValue(attribute, "position")
		'set value
		if tv.Value <> Cstr(attribute.Pos + 1) then
			'inform user
			Repository.WriteOutput outPutName, now() & " Correcting position tag for '"& element.Name & "." & attribute.name & "'", element.ElementID
			tv.Value = attribute.Pos + 1
			tv.Update
		end if
	next
end function

function addMissingPositionTags (element)
	'get attributes with missing position tags
	dim sqlGetAttributes
	sqlGetAttributes = " select a.ID from t_attribute a                      " & vbNewLine & _
						" inner join t_object o on o.Object_ID = a.Object_ID " & vbNewLine & _
						" where a.Object_ID = " & element.ElementID & "      " & vbNewLine & _
						" and (a.Stereotype = 'XSDelement'                   " & vbNewLine & _
						" or  o.Object_Type = 'Enumeration')                 " & vbNewLine & _
						" and not exists (select * from t_attributeTag tv    " & vbNewLine & _
						" 				where tv.ElementID = a.ID            " & vbNewLine & _
						" 				and tv.Property = 'position')        "
	dim attributes
	set attributes = getAttributesFromQuery(sqlGetAttributes)
	dim attribute as EA.Attribute
	for each attribute in attributes
		'inform user
		Repository.WriteOutput outPutName, now() & " Add position tag for '"& element.Name & "." & attribute.name & "'", element.ElementID
		'add position tag
		dim tv as EA.TaggedValue
		set tv = attribute.TaggedValues.AddNew("position", attribute.Pos + 1)
		tv.update
	next
end function

function correctAttributePositions(element)
	dim sqlGetAttributeInOrder
	sqlGetAttributeInOrder = "select a.ID from t_attribute a                 " & vbNewLine & _
							" where a.Object_ID = " & element.ElementID & "  " & vbNewLine & _
							" order by a.pos, a.Name                         "
	dim attributes
	set attributes = getAttributesFromQuery(sqlGetAttributeInOrder)
	dim attribute as EA.Attribute
	dim i
	i = 0
	for each attribute in attributes
		if attribute.Pos <> i then
			'inform user
			Repository.WriteOutput outPutName, now() & " Correct position for '"& element.Name & "." & attribute.name & "'", element.ElementID
			attribute.Pos = i
			attribute.Update
		end if
		'up the counter
		i = i + 1
	next
end function

main