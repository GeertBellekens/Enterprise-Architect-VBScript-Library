'[path=\Projects\Project B\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Correct XSD Datatypes
' Author: Geert Bellekens
' Purpose: Set a reference to the XSD Datatypes for XSDSimpleTypes and XSDelements
' Date: 2023-09-01
'

const outPutName = "Correct XSD datatypes"
const XSDdatatypesPackageGUID = "{9047E8CB-6D6A-47ec-82B9-16FA22D288D1}"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get the selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	'let the user know we started
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName & " for package '"& package.Name &"'", 0
	'do the actual work
	correctXSDDatatypes package
	'let the user know it is finished
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName & " for package '"& package.Name &"'", 0
end sub

function correctXSDDatatypes(package)
	'get XSDDatatypes dictionary
	dim XSDDatatypes
	set XSDDatatypes = getXSDDatatypes()
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	'elements
	processXSDelements packageTreeIDString, XSDDatatypes
	'simple types
	processXSDSimpleTypes packageTreeIDString, XSDDatatypes
	
end function

function getXSDDatatypes
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o                       " & vbNewLine & _
				" inner join t_package p on p.Package_ID = o.Package_ID   " & vbNewLine & _
				" where o.Stereotype = 'XSDsimpleType'                    " & vbNewLine & _
				" and p.ea_guid = '"& XSDdatatypesPackageGUID &"'         "
	dim result
	set result = getElementDictionaryFromQuery(sqlGetData)
	'return
	set getXSDDatatypes = result
end function

function processXSDelements(packageTreeIDString, XSDDatatypes)
	dim sqlGetData
	sqlGetData = "select a.ID                                          " & vbNewLine & _
				" from t_attribute a                                   " & vbNewLine & _
				" inner join t_object o on o.Object_ID = a.Object_ID   " & vbNewLine & _
				" where a.Stereotype = 'XSDelement'                    " & vbNewLine & _
				" and a.Classifier = 0                                 " & vbNewLine & _
				" and o.Package_ID in (" & packageTreeIDString &")     "
	dim attributes
	set attributes = getAttributesFromQuery(sqlGetData)
	dim attribute as EA.Attribute
	dim owner as EA.Element
	set owner = nothing
	for each attribute in Attributes
		'check if the same as previous owner
		if not owner is nothing then
			if not owner.ElementID = attribute.ParentID then
				set owner = nothing
			end if
		end if
		'get owner if needed
		set owner = Repository.GetElementByID(attribute.ParentID)
		'inform user
		Repository.WriteOutput outPutName, now() & " Processing '" & owner.Name & "."& attribute.Name &"'", 0
		if XSDDatatypes.Exists(attribute.Type) then
			dim datatypeElement as EA.Element
			set datatypeElement = XSDDatatypes(attribute.Type)
			attribute.ClassifierID = datatypeElement.ElementID
			attribute.Update
		else	
			'inform user
			Repository.WriteOutput outPutName, now() & " ERROR: Datatype '" & attribute.Type & "' on '" & owner.Name & "."& attribute.Name &"' not found in XSDDatatypes", 0
		end if
	next
end function

function processXSDSimpleTypes(packageTreeIDString, XSDDatatypes)
	'TODO
	'get XSDSimpleTypes that have "Parent=string;" in the Genlinks column, and no generalization to an XSDDatatype
	'remove the genlinks, and replace by generalization
end function

main