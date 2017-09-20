'[path=\Projects\Project K\Message Composer Scripts]
'[group=Message Composer Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Transform Message
' Author: Geert Bellekens
' Purpose: script to be executed as custom script after generating a message using the EA Message Composer.
' 		   this script will update all values of the tagged value "Mogelijk geen waarde" to the value "Ja" in the just created/updated message.
' Date: '2017-03-21
'
sub main
	' get the message package
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetCurrentDiagram
	if not currentDiagram is nothing then
		dim messagePackage as EA.Package
		set messagePackage = Repository.GetPackageByID(currentDiagram.PackageID)
		' ask the user for confirmation before updating the tags just in case something went wrong with the generation of the subset and we have the wrong diagram
		dim response
		response = msgbox("Update all 'Mogelijk geen waarde' of the subset in '" & messagePackage.Name & "' to the value 'Ja'?", vbYesNo+vbQuestion, "Update 'Mogelijk geen waarde' tags?")
		if response = vbYes then
			' get the package ID's of the package branch to be used in the SQL update statement
			dim packageIDTreeString 
			packageIDTreeString = getPackageTreeIDString(messagePackage)
			' update the tagged values using an SQL update statement
			dim updateAttributeTagsSQL
			updateAttributeTagsSQL = "update t_attributetag set [Value] = 'Ja'                     " & _
									" where [Property] = 'Mogelijk geen waarde'                    " & _
									" and [ElementID] in                                           " & _
									" (                                                            " & _
									"  select a.ID from t_attribute a                              " & _
									"  inner join t_object o on a.Object_ID = o.Object_ID      " & _
									"  where o.[Package_ID] in (" & packageIDTreeString & ")       " & _
									" )                                                            "
			'execute the update
			Repository.Execute updateAttributeTagsSQL
			dim updateConnectortagsSQL
			updateConnectortagsSQL = "update t_connectorTag set [Value] = 'Ja'                     " & _
									" where [Property] = 'Mogelijk geen waarde'                    " & _
									" and [ElementID] in                                           " & _
									" (                                                            " & _
									"  select c.[Connector_ID] from t_connector c                  " & _
									"  inner join t_object o on c.[Start_Object_ID] = o.Object_ID" & _
									"  where o.[Package_ID] in (" & packageIDTreeString & ")       " & _
									" )                                                            "
			'execute the update
			Repository.Execute updateConnectortagsSQL
		end if
	end if
	
end sub

'get the package id string of the given package tree
function getPackageTreeIDString(package)
	'initialize at "0"
	getPackageTreeIDString = "0"
	dim packageTree
	dim currentPackage as EA.Package
	if not package is nothing then
		'get the whole tree of the selected package
		set packageTree = getPackageTree(package)
		' get the id string of the tree
		getPackageTreeIDString = makePackageIDString(packageTree)
	end if 
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


main