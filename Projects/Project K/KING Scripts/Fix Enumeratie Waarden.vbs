'[path=\Projects\Project K\KING Scripts]
'[group=KING Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Fix Enumeratie Waarden
' Author: Geert Bellekens
' Purpose: the name of enumeration values in the different models should remain the same in the different models SIM, UGM, BSM.
'         This script will check if the name of the enumeration value is different and make it the same for the selected package and everthing underneath.
' Date: 2017-03-11
'

const outputTabName = "Fix Enumeratie Waarden"

sub main
	'setup output
	Repository.CreateOutputTab outputTabName
	Repository.ClearOutput outputTabName
	Repository.EnsureOutputVisible outputTabName
	'disable gui updates
	Repository.EnableUIUpdates = 0
	'tell the user we are starting
	Repository.WriteOutput outputTabName, now() & ": Starting Fixing Enumeratie Waarden",0	
	'select all the tagged values in the current 
	'first get the currently package id tree string
	dim selectedPackageIDString
	selectedPackageIDString = getSelectedPackageIDString()
	'fix enumeration values
    fixEnumerationValues selectedPackageIDString             
	'disable gui updates
	Repository.EnableUIUpdates = 1
	'reload 
	Repository.RefreshModelView(0)
	'tell the user we are finished
	Repository.WriteOutput outputTabName, now() & ": Finished Fixing Enumeratie Waarden",0	
end sub

function fixEnumerationValues(selectedPackageIDString)
	dim getEnumerationValues
	getEnumerationValues =  "select a.[ID],ao.name from (((t_attribute a                                                               " & _
							" inner join t_object o on o.Object_ID = a.Object_ID)                                                  " & _
							" inner join [t_attributetag] tv on tv.[ElementID] = a.[ID])                                               " & _
							" inner join [t_attribute] ao on ao.[ea_guid] like tv.VALUE)                                               " & _
							" where (o.[Object_Type] = 'Enumeration' or (o.[Object_Type] = 'Class' and o.[Stereotype] = 'Enumeration'))" & _
							" and a.[Name] <> ao.[Name]                                                                                " & _
							" and o.Package_ID in ("& getSelectedPackageIDString &")       			 								   "    						       				
	dim xmlResult
	xmlResult = Repository.SQLQuery(getEnumerationValues)
	dim enumValues
	enumValues = convertQueryResultToArray(xmlResult)
	dim i
	for i = Lbound(enumValues) to Ubound(enumValues) -1 step 1
		'get the values from the array
		dim attributeID
		dim originalName
		attributeID = enumValues(i,0)
		originalName = enumValues(i,1)
		'get the attribute object
		dim attribute as EA.Attribute
		set attribute = Repository.GetAttributeByID(attributeID)
		if not attribute is nothing _
		and attribute.Name <> originalName then
			'get the owner of th attribute for reporting
			dim owner as EA.Element
			set owner = Repository.GetElementByID(attribute.ParentID)
			'tell the user we are updating an attribute
			Repository.WriteOutput outputTabName, now() & " Updating value '" & owner.name & "." & attribute.Name & "' from '" &_
												attribute.Name & "' to '" & originalName & "'" ,owner.ElementID
			attribute.Name = originalName 
			attribute.Update
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