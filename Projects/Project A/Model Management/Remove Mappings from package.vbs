'[path=\Projects\Project A\Model Management]
'[group=Model Management]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Remove Mappings from package
' Author: Geert Bellekens
' Purpose: Remove all mapping tagged values from the selected package
' Date: 2019-12-04
'
const outPutName = "Remove mappings"
const elementMappingTag = "sourceElement"
const attributeMappingTag = "linkedAttribute"
const associationMappingTag = "linkedAssociation"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get selected package
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage
	if not selectedPackage is nothing then
		'ask for confirmation
		dim userIsSure
		userIsSure = Msgbox("Are you sure you want to remove the mappings from '" &selectedPackage.Name & "' and all its subPackages?'" & vbNewLine _
							& "Press 'Yes' to remove invalid mappings only" & vbNewLine _
							& "Press 'No' to remove ALL mappings" _
							, vbYesNoCancel+vbExclamation, "Remove mappings from " &selectedPackage.Name & "?")
		if userIsSure = vbYes or userIsSure = vbNo then
			Repository.WriteOutput outPutName, now() & " Starting remove mappings for package '"& selectedPackage.Name &"'", 0
			'remove the mappings
			dim allMappings
			if userIsSure = vbYes then
				allMappings = false
			else
				allMappings = true
			end if
			'remove the mappings
			removeMappings selectedPackage, allMappings
			'let user know
			Repository.WriteOutput outPutName, now() & " Finished remove mappings for package '"& selectedPackage.Name &"'", 0
		end if
	end if
end sub

function removeMappings(selectedPackage, allMappings)
	'get all items with mappings
	dim mappingItems
	set mappingItems = getAllmappingItems(selectedPackage, allMappings)
	dim mappingItem
	dim i 
	i = 0
	for each mappingItem in mappingItems
		i = i  + 1
		'let user know
		Repository.WriteOutput outPutName, now() & " Index: " & i &  " Removing mappings for item '"& mappingItem.Name &"'", 0
		removeMappingsFromItem mappingItem, allMappings
	next
end function

function removeMappingsFromItem(mappingItem, allMappings)
	dim i
	dim taggedValue as EA.TaggedValue
	for i = mappingItem.TaggedValues.Count -1 to 0 step - 1
		dim delete
		delete = false 'initialize
		set taggedValue = mappingItem.TaggedValues.GetAt(i)
		select case taggedValue.Name
			case elementMappingTag
				if allMappings then
					'remove all mappings
					delete = true
				else
					dim targetElement as EA.Element
					set targetElement = Repository.GetElementByGuid(taggedValue.Value)
					if targetElement is nothing then
						'remove only invalid mappings
						delete = true
					end if
				end if
			case attributeMappingTag
				if allMappings then
					'remove all mappings
					delete = true
				else
					dim targetAttribute as EA.Attribute
					set targetAttribute = Repository.GetAttributeByGuid(taggedValue.Value)
					if targetAttribute is nothing then
						'remove only invalid mappings
						delete = true
					end if
				end if
			case associationMappingTag
				if allMappings then
					'remove all mappings
					delete = true
				else
					dim targetConnector as EA.Connector
					set targetConnector = Repository.GetAttributeByGuid(taggedValue.Value)
					if targetConnector is nothing then
						'remove only invalid mappings
						delete = true
					end if
				end if
		end select
		'actually delete the mapping tagged value
		if delete then
			'inform user
			Repository.WriteOutput outPutName, now() & " Deleting tag '" & taggedValue.Name & "' with value '" & taggedValue.Value & "'" , 0
			mappingItem.TaggedValues.DeleteAt i, false
		end if
	next
end function

function getAllmappingItems(selectedPackage, allMappings)
	'inform user
	Repository.WriteOutput outPutName, now() & " Getting mapping items", 0
	'get package Tree ids
	dim packageTreeIDs 
	packageTreeIDs = getPackageTreeIDString(selectedPackage)
	dim mappingItems
	set mappingItems = CreateObject("System.Collections.ArrayList")
	'get mapping Elements
	dim sqlGetMappingElements
	sqlGetMappingElements = "select mp.* from ( 								       								" & vbNewLine & _
							" select tv.Object_ID from t_objectproperties tv								    	" & vbNewLine & _
							" where tv.Property = '" & elementMappingTag & "'										" & vbNewLine
	if not allMappings then 
		sqlGetMappingElements =  sqlGetMappingElements & _
						    " and not exists (select o.ea_guid from t_object o where o.ea_guid = tv.Value)		   " & vbNewLine
	end if
	sqlGetMappingElements =  sqlGetMappingElements & _						
							" union																				   " & vbNewLine & _
							" select tv.Object_ID from t_objectproperties tv									   " & vbNewLine & _
							" where tv.Property = '" & attributeMappingTag & "'									   " & vbNewLine
	if not allMappings then 
		sqlGetMappingElements =  sqlGetMappingElements & _
							" and not exists (select a.ID from t_attribute a where a.ea_guid = tv.Value)		   " & vbNewLine
	end if
	sqlGetMappingElements =  sqlGetMappingElements & _	
							" union																				   " & vbNewLine & _
							" select tv.Object_ID from t_objectproperties tv									   " & vbNewLine & _
							" where tv.Property = '" & associationMappingTag & "'		     					   " & vbNewLine
	if not allMappings then 
		sqlGetMappingElements =  sqlGetMappingElements & _
							" and not exists (select c.Connector_ID from t_connector c where c.ea_guid = tv.Value) " & vbNewLine
	end if
	'add restriction on package tree ids
	sqlGetMappingElements =  sqlGetMappingElements & _
							" ) mp                                                    								" & vbNewLine & _
							" inner join t_object o on mp.Object_ID = o.Object_ID  									" & vbNewLine & _
							" where o.Package_ID in (" & packageTreeIDs & ")        								"
	
	'debug
	'writeFile "H:\temp\elementQuery.txt", sqlGetMappingElements
	'get the elements
	set mappingItems = getElementsFromQuery(sqlGetMappingElements)
	'inform user
	Repository.WriteOutput outPutName, now() & " Mapping elements found: " & mappingItems.count, 0
	'get mapping attributes
	dim sqlGetMappingAttributes
	sqlGetMappingAttributes = "select mp.* from ( 								       								" & vbNewLine & _
							" select tv.ElementID from t_attributetag tv											" & vbNewLine & _
							" where tv.Property = '" & elementMappingTag & "'										" & vbNewLine
	if not allMappings then 
		sqlGetMappingAttributes =  sqlGetMappingAttributes & _
						    " and not exists (select o.ea_guid from t_object o where o.ea_guid = tv.Value)		   " & vbNewLine
	end if
	sqlGetMappingAttributes =  sqlGetMappingAttributes & _						
							" union																					" & vbNewLine & _
							" select tv.ElementID from t_attributetag tv											" & vbNewLine & _
							" where tv.Property = '" & attributeMappingTag & "'										" & vbNewLine
	if not allMappings then 
		sqlGetMappingAttributes =  sqlGetMappingAttributes & _
							" and not exists (select a.ID from t_attribute a where a.ea_guid = tv.Value)			" & vbNewLine
	end if
	sqlGetMappingAttributes =  sqlGetMappingAttributes & _	
							" union																				 	" & vbNewLine & _
							" select tv.ElementID from t_attributetag tv											" & vbNewLine & _
							" where tv.Property = '" & associationMappingTag & "'		     						" & vbNewLine
	if not allMappings then 
		sqlGetMappingAttributes =  sqlGetMappingAttributes & _
							" and not exists (select c.Connector_ID from t_connector c where c.ea_guid = tv.Value) " & vbNewLine
	end if
	'add restriction on package tree ids
	sqlGetMappingAttributes =  sqlGetMappingAttributes & _
							" ) mp                                                    								" & vbNewLine & _
							" inner join t_attribute a on a.ID = mp.ElementID	 									" & vbNewLine & _
							" inner join t_object o on a.Object_ID = o.Object_ID 									" & vbNewLine & _
							" where o.Package_ID in (" & packageTreeIDs & ")        								"  
	'debug
	writeFile "H:\temp\attributeQuery.txt", sqlGetMappingAttributes
	'get the attributes from the query
	dim mappingAttributes
	set mappingAttributes = getAttributesFromQuery(sqlGetMappingAttributes)
	'add attributes to mapping items
	mappingItems.AddRange(mappingAttributes)
	'inform user
	Repository.WriteOutput outPutName, now() & " Mapping attributes found: " & mappingAttributes.count, 0
	'get mapping associations
	dim sqlGetMappingAssociations
	sqlGetMappingAssociations = "select mp.* from ( 								       							" & vbNewLine & _
							" select tv.ElementID from t_connectortag tv											" & vbNewLine & _
							" where tv.Property = '" & elementMappingTag & "'										" & vbNewLine
	if not allMappings then 
		sqlGetMappingAssociations =  sqlGetMappingAssociations & _
						    " and not exists (select o.ea_guid from t_object o where o.ea_guid = tv.Value)		 	" & vbNewLine
	end if
	sqlGetMappingAssociations =  sqlGetMappingAssociations & _						
							" union																					" & vbNewLine & _
							" select tv.ElementID from t_connectortag tv											" & vbNewLine & _
							" where tv.Property = '" & attributeMappingTag & "'										" & vbNewLine
	if not allMappings then 
		sqlGetMappingAssociations =  sqlGetMappingAssociations & _
							" and not exists (select a.ID from t_attribute a where a.ea_guid = tv.Value)		 	" & vbNewLine
	end if
	sqlGetMappingAssociations =  sqlGetMappingAssociations & _	
							" union																					" & vbNewLine & _
							" select tv.ElementID from t_connectortag tv											" & vbNewLine & _
							" where tv.Property = '" & associationMappingTag & "'		     						" & vbNewLine
	if not allMappings then 
		sqlGetMappingAssociations =  sqlGetMappingAssociations & _
							" and not exists (select c.Connector_ID from t_connector c where c.ea_guid = tv.Value) " & vbNewLine
	end if
	'add restriction on package tree ids
	sqlGetMappingAssociations =  sqlGetMappingAssociations & _
							" ) mp                                                    								" & vbNewLine & _
							" inner join t_connector c on c.Connector_ID = mp.ElementID	 							" & vbNewLine & _
							" inner join t_object o on c.Start_Object_ID = o.Object_ID 								" & vbNewLine & _
							" where o.Package_ID in (" & packageTreeIDs & ")        								"  
	'debug
	'writeFile "H:\temp\associationQuery.txt", sqlGetMappingAttributes
	'get the associations
	dim mappingAssociations
	set mappingAssociations = getConnectorsFromQuery(sqlGetMappingAssociations) 
	'add the associations
	mappingItems.AddRange(mappingAssociations)
	'inform user
	Repository.WriteOutput outPutName, now() & " Mapping attributes found: " & mappingAssociations.count, 0
	'return
	set getAllmappingItems = mappingItems
end function



main