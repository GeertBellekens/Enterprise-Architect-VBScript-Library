'[path=\Projects\Project K\KING Scripts]
'[group=KING Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Fix Indicatie Kerngegeven
' Author: Geert Bellekens
' Purpose: fix duplicate instances of the tagged value Indicatie kerngegeven in order to only keep one per attribute (or association)
' the rules to folow are
'  - if there are different values in the different tagged values then "Ja" should be used
'  - if all the values are the same then that value should be used
'  - all other tagged values with the same name should be removed 
' Date: 2017-03-11
'

const outputTabName = "Fix Indicatie Kerngegeven"

sub main
	'setup output
	Repository.CreateOutputTab outputTabName
	Repository.ClearOutput outputTabName
	Repository.EnsureOutputVisible outputTabName
	'tell the user we are starting
	Repository.WriteOutput outputTabName, now() & ": Starting Fixing Indicatie Kerngegeven",0	
	'select all the tagged values in the current 
	'first get the currently package id tree string
	dim selectedPackageIDString
	selectedPackageIDString = getSelectedPackageIDString()
	'fix attribute tags
    fixAttributeTags selectedPackageIDString             
	'fix connector tags
	fixConnectorTags selectedPackageIDString
	'tell the user we are finished
	Repository.WriteOutput outputTabName, now() & ": Finished Fixing Indicatie Kerngegeven",0	
end sub

function fixAttributeTags(getSelectedPackageIDString)
	dim getAttributesSQL
	getAttributesSQL =  "select distinct a.ID from (((t_attributetag tv              			   " & _
						" inner join t_attributetag tv2 on tv2.ElementID = tv.ElementID)           " & _
						" inner join t_attribute a on tv.ElementID = a.ID)                         " & _
						" inner join t_object o on o.Object_ID = a.Object_ID)                      " & _
						" where (tv.Property = 'Indicatie kerngegeven'                             " & _
						" or tv.Property = 'Indicate kerngegeven')                                " & _
						" and o.Package_ID in ("& getSelectedPackageIDString &")       			   " & _
						" and (tv2.Property = 'Indicatie kerngegeven'  							   " & _
						" or tv2.Property = 'Indicate kerngegeven')                                " & _
						" and tv2.ea_guid <> tv.ea_guid             						       "					
	dim attributes
	set attributes = getattributesFromQuery(getAttributesSQL)
	dim attribute as EA.Attribute
	for each attribute in attributes
		Repository.WriteOutput outputTabName, now() & ": Processing attribute '" & attribute.Name & "'"  ,attribute.ParentID	
		fixKernGegevenTags attribute
	next
end function

function fixConnectorTags(getSelectedPackageIDString)
	dim getConnectorsSQL
	getConnectorsSQL =  "select distinct c.[Connector_ID]                                " & _
						" from (((t_connectortag tv              			             " & _
						" inner join t_connectortag tv2 on tv2.ElementID = tv.ElementID) " & _
						" inner join t_connector c on tv.[ElementID] = c.[Connector_ID]) " & _
						" inner join t_object ost on ost.Object_ID = c.[Start_Object_ID])" & _
						" where (tv.Property = 'Indicatie kerngegeven'                   " & _
						" or tv.Property = 'Indicate kerngegeven')                       " & _
						" and ost.Package_ID in ("& getSelectedPackageIDString &")       " & _
						" and (tv2.Property = 'Indicatie kerngegeven'  					 " & _
						" or tv2.Property = 'Indicate kerngegeven')                      " & _
						" and tv2.ea_guid <> tv.ea_guid                                  "				
	dim connectors
	set connectors = getconnectorsFromQuery(getConnectorsSQL)
	dim connector as EA.Connector
	for each connector in connectors
		Repository.WriteOutput outputTabName, now() & ": Processing connector '" & connector.Name & "'"  ,connector.ClientID	
		fixKernGegevenTags connector
	next
end function


function fixKernGegevenTags(tagOwner)
	dim taggedValue as EA.TaggedValue
	'initialize tagged value
	set taggedValue = nothing
	dim currentTag as EA.TaggedValue
	dim tagValue
	'loop backward through the tagged values to keep only one
	'get all tagged values with the kerngegeven name and remember the one to keep
	dim kerngegevenTags
	set kerngegevenTags = CreateObject("System.Collections.ArrayList")
	for each currentTag in tagOwner.TaggedValues
		if currentTag.Name = "Indicatie kerngegeven" _
		OR currentTag.Name = "Indicate kerngegeven" then
			'if it has the correct name and is part of the profile
			'then keep remember this one as the one to keep
			if currentTag.Name = "Indicatie kerngegeven" _
			AND instr(currentTag.FQName,"::") > 0 _
			AND taggedValue is nothing then
				set taggedValue = currentTag
			end if
		'add the tag to the list
		kerngegevenTags.Add currentTag
		end if
	next
	'check if the we one to keep. If not take the first one and update its name as well
	if taggedValue is nothing then
		set taggedValue = kerngegevenTags(0)
	end if
	'first determine the value kerngegeven
	tagValue = getValueForKerngegeven(kerngegevenTags)
	'set the value for the kerngegeven, only update if needed
	if taggedValue.Value <> tagValue _
	OR taggedValue.Name <> "Indicatie kerngegeven" then
		'tell the user we are updating a tag
		Repository.WriteOutput outputTabName, now() & ": updating tag '" & taggedValue.Name & "' to name 'Indicatie Kerngegeven' and value '" & tagValue & "'"  ,0	
		taggedValue.Name = "Indicatie kerngegeven"
		taggedValue.Value = tagValue
		taggedValue.Update
	end if
	'now delete all the others
	dim i
	for i = tagOwner.TaggedValues.Count -1 to 0 step -1
		'get the current tag
		set currentTag = tagOwner.TaggedValues(i)
		'check if not the one we want to keep
		if currentTag.TagGUID <> taggedValue.TagGUID then
			'if the name matches then delete it
			if currentTag.Name = "Indicatie kerngegeven" _
			OR currentTag.Name = "Indicate kerngegeven" then
				'tell the user we are deleting a tag
				Repository.WriteOutput outputTabName, now() & ": Removing tag '" & currentTag.Name & "'" ,0	
				'delete the duplicate
				tagOwner.TaggedValues.DeleteAt i, false 
			end if
		end if
	next
end function

function getValueForKerngegeven(kerngegevenTags)
	getValueForKerngegeven = "" '(default value)
	'first determine the value the tagged value should have
	dim taggedValue
	for each taggedValue in kerngegevenTags
		if getValueForKerngegeven = "" then
			'set the vaue for the first time
			getValueForKerngegeven = taggedValue.Value
		elseif getValueForKerngegeven <> taggedValue.Value AND taggedValue.Value <> "" then
			'difference found so default value = 'Ja'
			getValueForKerngegeven = "Ja"
		end if
	next
end function

'make an id string out of the package ID of the given packages
function makePackageIDString(packages)
	dim package as EA.Package
	dim idString
	idString = ""
	dim addComma 
	addComma = false
	for each package in packages
		if addComma then
			idString = idString & ","
		else
			addComma = true
		end if
		idString = idString & package.PackageID
	next 
	'if there are no packages then we return "0"
	if packages.Count = 0 then
		idString = "0"
	end if
	'return idString
	makePackageIDString = idString
end function

'returns an ArrayList of the given package and all its subpackages recursively
function getPackageTree(package)
	dim packageList
	set packageList = CreateObject("System.Collections.ArrayList")
	addPackagesToList package, packageList
	set getPackageTree = packageList
end function

'add the given package and all subPackges to the list (recursively
function addPackagesToList(package, packageList)
	dim subPackage as EA.Package
	'add the package itself
	packageList.Add package
	'add subpackages
	for each subPackage in package.Packages
		addPackagesToList subPackage, packageList
	next
end function

function getSelectedPackageIDString()
	'get IDString the currently selected package in the project project (recursively)
	dim selectedPackage as EA.Package
	'initialize
	set selectedPackage = nothing
	set selectedPackage = Repository.GetTreeSelectedPackage()
	dim packageTree
	set packageTree = getPackageTree(selectedPackage)
	'return the package ID string
	getSelectedPackageIDString = makePackageIDString(packageTree)
end function

function getConnectorsFromQuery(sqlQuery)
	dim xmlResult
	xmlResult = Repository.SQLQuery(sqlQuery)
	dim connectorIDs
	connectorIDs = convertQueryResultToArray(xmlResult)
	dim connectors 
	set connectors = CreateObject("System.Collections.ArrayList")
	dim connectorID
	dim connector as EA.Connector
	for each connectorID in connectorIDs
		if connectorID > 0 then
			set connector = Repository.GetConnectorByID(connectorID)
			if not connector is nothing then
				connectors.Add(connector)
			end if
		end if
	next
	set getConnectorsFromQuery = connectors
end function

function getattributesFromQuery(sqlQuery)
	dim xmlResult
	xmlResult = Repository.SQLQuery(sqlQuery)
	dim attributeIDs
	attributeIDs = convertQueryResultToArray(xmlResult)
	dim attributes 
	set attributes = CreateObject("System.Collections.ArrayList")
	dim attributeID
	dim attribute as EA.Attribute
	for each attributeID in attributeIDs
		if attributeID > 0 then
			set attribute = Repository.GetAttributeByID(attributeID)
			if not attribute is nothing then
				attributes.Add(attribute)
			end if
		end if
	next
	set getattributesFromQuery = attributes
end function

'converts the query results from Repository.SQLQuery from xml format to a two dimensional array of strings
Public Function convertQueryResultToArray(xmlQueryResult)
    Dim arrayCreated
    Dim i 
    i = 0
    Dim j 
    j = 0
    Dim result()
    Dim xDoc 
    Set xDoc = CreateObject( "MSXML2.DOMDocument" )
    'load the resultset in the xml document
    If xDoc.LoadXML(xmlQueryResult) Then        
		'select the rows
		Dim rowList
		Set rowList = xDoc.SelectNodes("//Row")

		Dim rowNode 
		Dim fieldNode
		arrayCreated = False
		'loop rows and find fields
		For Each rowNode In rowList
			j = 0
			If (rowNode.HasChildNodes) Then
				'redim array (only once)
				If Not arrayCreated Then
					ReDim result(rowList.Length, rowNode.ChildNodes.Length)
					arrayCreated = True
				End If
				For Each fieldNode In rowNode.ChildNodes
					'write f
					result(i, j) = fieldNode.Text
					j = j + 1
				Next
			End If
			i = i + 1
		Next
	end if
    convertQueryResultToArray = result
End Function

main