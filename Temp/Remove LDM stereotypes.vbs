'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Remove LDM Stereotypes
' Author: Geert Bellekens
' Purpose: Remove all the LDM stereotypes in the model
' Date: 2019-05-22
'

'name of the output tab
const outPutName = "Remove LDM stereotypes"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp for start
	Repository.WriteOutput outPutName,now() & " Starting removing LDM stereotypes"  , 0
	'get selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage
	if package is nothing then
		exit sub
	end if
	'remove the stereotypes
	removeLDMStereotypes(package)
	'set timestamp for end
	Repository.WriteOutput outPutName,now() & " Finished removing LDM stereotypes"  , 0
end sub

function removeLDMStereotypes(package)
	'get the package tree id's
	dim packageTreeIDs
	packageTreeIDs = getPackageTreeIDString(package)
	'remove from elements
	removeLDMStereotypeOnElements packageTreeIDs
	'remove from attributes
	removeLDMStereotypeOnAttributes packageTreeIDs
	'remove from connectors
	removeLDMStereotypeOnConnectors packageTreeIDs
end function

function removeLDMStereotypeOnElements(packageTreeIDs)
	'get all elements
	dim sqlGetLdmElements 
	sqlGetLdmElements = "select o.Object_ID from t_object o where o.stereotype like 'LDM%'" & vbNewLine & _
						" and o.Package_ID in (" & packageTreeIDs & ")        "
	dim ldmElements
	set ldmElements = getElementsFromQuery(sqlGetLdmElements)
	dim ldmElement as EA.Element
	for each ldmElement in ldmElements
		Repository.WriteOutput outPutName,now() & " Removing stereotype from element '" & ldmElement.Name & "'"  , 0
		ldmElement.StereotypeEx = ""
		ldmElement.Update
	next
end function

function removeLDMStereotypeOnAttributes(packageTreeIDs)
	'get all Attributes
	dim sqlGetLdmAttributes 
	sqlGetLdmAttributes =   "select a.ID from t_attribute a                       " & vbNewLine & _
							" inner join t_object o on o.Object_ID = a.Object_ID  " & vbNewLine & _
							" where a.stereotype like 'LDM%'                    " & vbNewLine & _
							" and o.Package_ID in (" & packageTreeIDs & ")        "
	dim ldmAttributes
	set ldmAttributes = getAttributesFromQuery(sqlGetLdmAttributes)
	dim ldmAttribute as EA.Attribute
	for each ldmAttribute in ldmAttributes
		Repository.WriteOutput outPutName,now() & " Removing stereotype from attribute '" & ldmAttribute.Name & "'"  , 0
		ldmAttribute.StereotypeEx = ""
		ldmAttribute.Update
	next
end function

function removeLDMStereotypeOnConnectors(packageTreeIDs)
	'get all Connectors
	dim sqlGetLdmConnectors 
	sqlGetLdmConnectors =  "select c.Connector_ID from t_Connector c                   " & vbNewLine & _
							" inner join t_object o on o.Object_ID = c.Start_Object_ID " & vbNewLine & _
							" where c.stereotype like 'LDM%'                           " & vbNewLine & _
							" and o.Package_ID in (" & packageTreeIDs & ")             "
	dim ldmConnectors
	set ldmConnectors = getConnectorsFromQuery(sqlGetLdmConnectors)
	dim ldmConnector as EA.Connector
	for each ldmConnector in ldmConnectors
		Repository.WriteOutput outPutName,now() & " Removing stereotype from connector '" & ldmConnector.Name & "'"  , 0
		ldmConnector.StereotypeEx = ""
		ldmConnector.Update
	next
end function

main