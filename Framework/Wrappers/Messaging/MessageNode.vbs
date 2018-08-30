'[path=\Framework\Wrappers\Messaging]
'[group=Messaging]

!INC Utils.Include
!INC Local Scripts.EAConstants-VBScript
' Author: Geert Bellekens
' Purpose: A wrapper class for a message node in a messaging structure
' Date: 2017-03-14

Class MessageNode 
	'private variables
	Private m_Name
	Private m_TypeElement
	Private m_TypeName
	Private m_Multiplicity
	Private m_ParentNode
	Private m_ChildNodes
	Private m_SourceAttribute
	Private m_SourceAssociationEnd
	Private m_SourceElement
	Private m_ValidationRules
	Private m_IsLeafNode
	Private m_CustomOrdering
	Private m_Order
	Private m_Facets
	Private m_MappedBusinessAttributes
	Private m_IncludeDetails
	Private m_BaseTypeName
	Private m_BaseTypeElement
	Private m_Message

	'constructor
	Private Sub Class_Initialize
		m_Name = ""
		set m_TypeElement = nothing
		m_TypeName = ""
		m_Multiplicity = ""
		set m_ParentNode = nothing
		set m_ChildNodes = CreateObject("System.Collections.ArrayList")
		set m_SourceAttribute = nothing
		set m_SourceAssociationEnd = nothing
		set m_SourceElement = nothing
		set m_ValidationRules = CreateObject("System.Collections.ArrayList")
		m_IsLeafNode = false
		m_CustomOrdering = false
		m_order = 0
		set m_Facets = CreateObject("Scripting.Dictionary")
		set m_MappedBusinessAttributes = CreateObject("System.Collections.ArrayList")
		m_IncludeDetails = false
		m_BaseTypeName = ""
		set m_BaseTypeElement = nothing
		set m_Message = nothing
	End Sub
	
	'public properties
	
	' IsLeafNode property.
	Public Property Get IsLeafNode
		IsLeafNode = m_IsLeafNode
	End Property
	
	' IncludeDetails property.
	Public Property Get IncludeDetails
		if not me.ParentNode is nothing then
			IncludeDetails = me.ParentNode.IncludeDetails
		else
			IncludeDetails = m_IncludeDetails
		end if
	End Property
	Public Property Let IncludeDetails(value)
		m_IncludeDetails = value
	End Property
	
	' Name property.
	Public Property Get Name
		Name = m_Name
	End Property
	Public Property Let Name(value)
		m_Name = value
	End Property
	
	' TypeElement property.
	Public Property Get TypeElement
		set TypeElement = m_TypeElement
	End Property
	Public Property Let TypeElement(value)
		set m_TypeElement = value
	End Property

	' ElementID property.
	Public Property Get ElementID
		if not me.TypeElement is nothing then
			ElementID = me.TypeElement.ElementID
		else
			ElementID = 0
		end if
	End Property

	' TypeName property
	Public Property Get TypeName
		if not me.TypeElement is nothing then
			TypeName = me.TypeElement.Name
		else
			TypeName = m_TypeName
		end if 
	End Property
	Public Property Let TypeName(value)
		m_TypeName = value
		'if the typename if different from the TypeElement name then we remove the type Element
		if not me.TypeElement is nothing then
			if value <> me.TypeElement.Name then
				me.TypeElement = nothing
			end if
		end if
	End Property
	' BaseTypeElement property.
	Public Property Get BaseTypeElement
		set BaseTypeElement = m_BaseTypeElement
	End Property
	Public Property Let BaseTypeElement(value)
		set m_BaseTypeElement = value
	End Property
	
	' BaseTypeName property.
	Public Property Get BaseTypeName
		if not me.BaseTypeElement is nothing then
			BaseTypeName = me.BaseTypeElement.Name
		else
			BaseTypeName = m_BaseTypeName
		end if 
	End Property
	Public Property Let BaseTypeName(value)
		m_BaseTypeName = value
		'if the typename if different from the TypeElement name then we remove the type Element
		if not me.BaseTypeElement is nothing then
			if value <> me.BaseTypeElement.Name then
				me.BaseTypeElement = nothing
			end if
		end if
	End Property
		
	' Multiplicity property.
	' only directly used if the source is element, else we use the Attribute or AssociationEnd multiplicity
	Public Property Get Multiplicity
		dim lower
		dim upper
		dim returnedMultiplicity
		if not me.SourceElement is nothing then
			returnedMultiplicity = m_Multiplicity
		elseif not me.sourceAttribute is nothing then
			returnedMultiplicity = determineMultiplicity(me.sourceAttribute.LowerBound,me.sourceAttribute.UpperBound, "1", "1")
		elseif not me.sourceAssociationEnd is nothing then
			returnedMultiplicity = sourceAssociationEnd.Cardinality
		end if
		'return the actual value
		Multiplicity = returnedMultiplicity
	End Property
	Public Property Let Multiplicity(value)
		if not me.SourceElement is nothing then
			m_Multiplicity = value
		end if
	End Property
	
	private function determineMultiplicity(lower,upper,defaultLower, defaultUpper)
		'check to make sur the values are filled in and replace them with the default values if not the case
		if len(lower) = 0 then
			lower = defaultLower
		end if
		if len(upper) = 0 then
			upper = defaultUpper
		end if
		'create the multiplicity string
		determineMultiplicity = lower & ".." & upper
	end function
	' ParentNode property.
	Public Property Get ParentNode
		set ParentNode = m_ParentNode
	End Property
	Public Property Let ParentNode(value)
		set m_ParentNode = value
	End Property

	' ChildNodes property.
	Public Property Get ChildNodes
		set ChildNodes = m_ChildNodes
	End Property
	Public Property Let ChildNodes(value)
		set m_ChildNodes = value
	End Property
	
	' SourceAttribute property.
	Public Property Get SourceAttribute
		set SourceAttribute = m_SourceAttribute
	End Property
	Public Property Let SourceAttribute(value)
		set m_SourceAttribute = value
	End Property

	' SourceAssociationEnd property.
	Public Property Get SourceAssociationEnd
		set SourceAssociationEnd = m_SourceAssociationEnd
	End Property
	Public Property Let SourceAssociationEnd(value)
		set m_SourceAssociationEnd = value
	End Property
	
	' SourceElement property.
	Public Property Get SourceElement
		set SourceElement = m_SourceElement
	End Property
	Public Property Let SourceElement(value)
		set m_SourceElement = value
	End Property
	
	' ValidationRules property.
	Public Property Get ValidationRules
		set ValidationRules = m_ValidationRules
	End Property
	Public Property Let ValidationRules(value)
		set m_ValidationRules = value
	End Property	
	
	'CustomOrdering property
	Public Property Get CustomOrdering
	  CustomOrdering = m_CustomOrdering
	End Property
	Public Property Let CustomOrdering(value)
		m_CustomOrdering = value
	End Property
	
	'Order property
	Public Property Get Order
	  Order = m_Order
	End Property
	Public Property Let Order(value)
		m_Order = value
	End Property
	
	' Facets property. (Dictionary with key = facet name, value = facet value)
	Public Property Get Facets
		set Facets = m_Facets
	End Property
	Public Property Let Facets(value)
		set m_Facets = value
	End Property
	
	' MappedBusinessAttributes property.
	Public Property Get MappedBusinessAttributes
		set MappedBusinessAttributes = m_MappedBusinessAttributes
	End Property
	Public Property Let MappedBusinessAttributes(value)
		set m_MappedBusinessAttributes = value
	End Property
	
	' Message property.
	Public Property Get Message
		set Message = m_Message
	End Property
	Public Property Let Message(value)
		set m_Message = value
	End Property
	
	Public function linkRuletoNode(validationRule, path)
		'initialize false
		linkRuletoNode = false
		if path.Count > 0 then
			if path.Count = 1 then
				'link it to this node
				m_ValidationRules.Add validationRule
				linkRuletoNode = true
			else
				'go deeper
				dim childNode
				for each childNode in me.ChildNodes
					dim newPath
					set newPath = nothing
					if path(1) = childNode.Name then
						if newPath is nothing then
							set newPath = CreateObject("System.Collections.ArrayList")
							'create new path removing the first part
							dim i
							for i = 1 to path.Count -1 step +1
								newPath.Add path(i)
								'return true
								linkRuletoNode = true
							next
						end if
						'go one level deeper
						linkRuletoNode = childNode.linkRuletoNode(validationRule, newPath)
					end if
				next
			end if
		end if
	end function
	
	'public functions
	public function intitializeWithSource(source,sourceConnector,in_multiplicity,in_validationRule,in_parentNode)
		'set validationRule
		if not in_validationRule is nothing then
			me.ValidationRule = in_validationRule
		end if
		'set parentNode
		if not in_parentNode is nothing then
			me.ParentNode = in_parentNode
		end if
		'check if source is Element, Atttribute, or AssociationEnd
		select case source.ObjectType
			case otElement
				setElementNode source,in_multiplicity
			case otAttribute
				setAttributeNode source
			case otConnectorEnd
				setConnectorEndNode source,sourceConnector
		end select
		Repository.WriteOutput outPutName, now() & " Processing node '" & me.Name & "'", 0
		'set the isLeafNode property
		setIsLeafNode
		'then load the child nodes
		if not me.IsLeafNode then
			loadChildNodes
		end if
	end function
	
	'set the source node in case the source is an element
	private function setElementNode(source,in_multiplicity)
		me.SourceElement = source
		me.Name = source.Name
		me.TypeElement = source
		me.Multiplicity = in_multiplicity
	end function
	
	'set the source in case of an attribute
	private function setAttributeNode(source)
		me.SourceAttribute = source
		'set the order
		me.Order = getSequencingKey(source)
		'set the name
		me.Name = source.Name
		'remove any underscores from the name in case of MIG6
		if me.CustomOrdering then
			me.Name = Replace(me.Name, "_","")
		end if
		'set the type
		dim attributeTypeObject as EA.Element
		set attributeTypeObject = nothing
		if source.ClassifierID > 0 then
			set attributeTypeObject = Repository.GetElementByID(source.ClassifierID)
			'if the attributeTypeObject is a «BDT» then we get the attribute with stereotype «CON» and name "content" and use it's type as the typeElement
			if attributeTypeObject.Stereotype = "BDT" then
				'get the content attribute
				dim conAttribute as EA.Attribute
				set conAttribute = nothing
				for each conAttribute in attributeTypeObject.Attributes
					if conAttribute.Stereotype = "CON" _
					  and conAttribute.Name = "content" then
						exit for
					end if
				next
				if not conAttribute is nothing then
					'get the facets from the conAttribute as well
					getFacets conAttribute
					if conAttribute.ClassifierID > 0 then
						dim conTypeObject as EA.Element
						set conTypeObject = Repository.GetElementByID(conAttribute.ClassifierID)
						'check for directXSD types
						me.BaseTypeElement = getBaseType(attributeTypeObject, conTypeObject)
						me.TypeElement = attributeTypeObject
					else
						me.TypeElement = attributeTypeObject
						me.BaseTypeName = conAttribute.Type
					end if
				else
					'content attribute not found, set error
					me.TypeName = "Error: BDT " & attributeTypeObject & " has no content attribute"
				end if
			else
				'regular attribute
				me.TypeElement = attributeTypeObject
				'find the parent class (not for enumerations)
				if not (me.TypeElement.Type = "Enumeration" _
					or lcase(me.TypeElement.Stereotype) = "enumeration") then
					dim parentClass as EA.Element
					for each parentClass in attributeTypeObject.BaseClasses
						me.BaseTypeElement = parentClass
						exit for 'return immediate
					next
				end if
			end if
		else
			me.TypeName = source.Type
		end if
		'get the facets
		getFacets source
		'set the mapped BusinessAttributes
		if me.CustomOrdering then 'only applicable for custom ordering
			dim taggedValue as EA.AttributeTag
			'find the tagged values with name mappedBusinessAttribute
			for each taggedValue in source.TaggedValues
				if lcase(taggedValue.Name) = "mappedbusinessattribute" then
					dim businessAttribute as EA.Attribute
					set businessAttribute = Repository.GetAttributeByGuid(taggedValue.Value)
					if not businessAttribute is nothing then
						MappedBusinessAttributes.Add businessAttribute
					end if
				end if
			next
		end if
	end function
	
	private function getBaseType(attributeTypeObject, conTypeObject)
		'initialize
		set getBaseType = conTypeObject
		'figure out of the attributeTypeObject has tagged value with name "directXSDType" and value "true"
		dim isDirectXSDType
		isDirectXSDType = false
		dim tv as EA.TaggedValue
		for each tv in attributeTypeObject.TaggedValues
			if lcase(tv.Name) = "directxsdtype" _
			and lcase(tv.Value) = "true" then
				isDirectXSDType = true
			end if
		next
		if isDirectXSDType then
			'find the parent class
			dim parentClass as EA.Element
			for each parentClass in attributeTypeObject.BaseClasses
				set getBaseType = parentClass
				exit for 'return immediate
			next
		end if
	end function
	
	'gets the facets from an attribute
	private function getFacets(sourceAttribute)
		dim tv as EA.TaggedValue
		'first loop the facets of the datatype
		dim datatype as EA.Element
		if sourceAttribute.ClassifierID > 0 then
			set datatype = Repository.GetElementByID(sourceAttribute.ClassifierID)
			for each tv in datatype.TaggedValues
				if len(tv.Value) > 0 _
				  and (tv.Name = "enumeration" _
				  or tv.Name = "fractionDigits" _
				  or tv.Name = "length" _
				  or tv.Name = "maxExclusive" _
				  or tv.Name = "maxInclusive" _
				  or tv.Name = "maxLength" _
				  or tv.Name = "minExclusive" _
				  or tv.Name = "minInclusive" _
				  or tv.Name = "minLength" _
				  or tv.Name = "pattern" _
				  or tv.Name = "totalDigits" _
				  or tv.Name = "whiteSpace") then
					me.Facets.Item(tv.Name) = tv.Value
				end if
			next
		end if
		'first loop the standard facets
		for each tv in sourceAttribute.TaggedValues
			if tv.Name = "enumeration" _
			  or tv.Name = "fractionDigits" _
			  or tv.Name = "length" _
			  or tv.Name = "maxExclusive" _
			  or tv.Name = "maxInclusive" _
			  or tv.Name = "maxLength" _
			  or tv.Name = "minExclusive" _
			  or tv.Name = "minInclusive" _
			  or tv.Name = "minLength" _
			  or tv.Name = "pattern" _
			  or tv.Name = "totalDigits" _
			  or tv.Name = "whiteSpace" then
				me.Facets.Item(tv.Name) = tv.Value
			end if
		next
		'then the overridden facets
		for each tv in sourceAttribute.TaggedValues
			if tv.Name = "override_enumeration" _
			  or tv.Name = "override_fractionDigits" _
			  or tv.Name = "override_length" _
			  or tv.Name = "override_maxExclusive" _
			  or tv.Name = "override_maxInclusive" _
			  or tv.Name = "override_maxLength" _
			  or tv.Name = "override_minExclusive" _
			  or tv.Name = "override_minInclusive" _
			  or tv.Name = "override_minLength" _
			  or tv.Name = "override_pattern" _
			  or tv.Name = "override_totalDigits" _
			  or tv.Name = "override_whiteSpace" then
				me.Facets.Item(Replace(tv.Name,"override_","")) = tv.Value
			end if
		next
	end function
	
	
	'set the source in case of a connectorEnd
	private function setConnectorEndNode(source,sourceConnector)
		me.SourceAssociationEnd = source
		'set the order
		me.Order = getSequencingKey(sourceConnector)
		dim endObject as EA.Element
		'get the end object 
		if source.End = "Supplier" then
			set endObject = Repository.GetElementByID(sourceConnector.SupplierID)
		else
			set endObject = Repository.GetElementByID(sourceConnector.ClientID)
		end if
		'set the name = name of role + name of end object + remove underscores
		if len(source.Role) > 0 then
			me.Name = source.Role & endObject.Name
			me.Name = Replace(me.Name, "_","")
		else
			'use the end object name as rolename
			me.Name = endObject.Name
		end if 
		'set the type
		me.TypeElement = endObject
	end function
	
	public function getSequencingKey(sourceItem)
		'initialize at 0
		getSequencingKey = 0
		dim taggedValue as EA.TaggedValue
		for each taggedValue in sourceItem.TaggedValues
			if Lcase(taggedValue.Name) = "sequencingkey" then
				on error resume next
				getSequencingKey = CInt(taggedValue.Value)
				if Err.Number <> 0 then
					err.Clear
				end if
				on error goto 0
				exit for
			end if
		next
	end function
	
	'Loads the child nodes for this message (resursively until we have reached all the leaves)
	public function loadChildNodes()
		'first remember the list of parent elements
		dim parents
		set parents = getParents(nothing)
		'TODO: load in correct order?
		'load attributes
		loadAllAttributeNodes parents 
		'load associations
		loadAllAssociationNodes parents
		'load nested classes?
		'reorder nodes
		if me.CustomOrdering then
			reOrderChildNodes
		end if
	end function
	
	public function reOrderChildNodes
		dim childNode
			dim i
		dim goAgain
		goAgain = false
		dim currentNode
		dim nextNode
		for i = 0 to me.ChildNodes.Count -2 step 1
			set currentNode = me.ChildNodes.Item(i)
			set nextNode = me.ChildNodes.Item(i +1)
			if  currentNode.Order > nextNode.Order then
				me.ChildNodes.RemoveAt(i +1)
				me.ChildNodes.Insert i, nextNode
				goAgain = true
			end if
		next
		'if we had to swap an element then we go over the list again
		if goAgain then
			reOrderChildNodes
		end if
	end function
	
	'gets the maximum depth of this node and add that to the given depth
	public function getDepth(in_depth)
		dim childNode
		dim maxDebth
		maxDebth = in_depth + 1
		for each childNode in me.ChildNodes
			dim currentDepth
			currentDepth = childNode.getDepth(in_depth +1)
			if currentDepth > maxDebth then
				maxDebth = currentDepth
			end if
		next
		getDepth = maxDebth
	end function
	
	'gets the output format for this node and its childnodes
	public function getOuput(current_order,currentPath,messageDepth, includeRules)
		'create the output
		dim nodeOutputList
		set nodeOutputList = CreateObject("System.Collections.ArrayList")
		dim currentNodeList
		'get the list for this node
		if me.ValidationRules.Count = 0 or not includeRules then
			set currentNodeList = getThisNodeOutput(current_order,currentPath, messageDepth,nothing, includeRules)
			'up or the order number
			current_order = current_order + 1
			'add the list for this node to the output
			nodeOutputList.Add currentNodeList
		else
			dim currentRule
			for each currentRule in me.ValidationRules
				set currentNodeList = getThisNodeOutput(current_order,currentPath, messageDepth,currentRule, includeRules)
				'up or the order number
				current_order = current_order + 1
				'add the list for this node to the output
				nodeOutputList.Add currentNodeList
			next
		end if
		'add this node to the currentPath
		dim mycurrentpath
		set myCurrentPath = CreateObject("System.Collections.ArrayList")
		myCurrentPath.AddRange(currentPath)
		myCurrentPath.Add me.Name
		'get the output for the child nodes
		dim childNode
		for each childNode in me.ChildNodes
			dim childOutPut
			set childOutPut = childNode.getOuput(current_order,myCurrentPath,messageDepth, includeRules)
			nodeOutputList.AddRange(childOutPut)
		next
		'return list
		set getOuput = nodeOutputList
	end function
	

	
	private function getThisNodeOutput(current_order,currentPath, messageDepth,validationRule, includeRules)
		'get the list for this node
		dim currentNodeList
		set currentNodeList = CreateObject("System.Collections.ArrayList")
		'add the order to the list
		currentNodeList.Add lpad(current_order,4,"0")
		'add the current Path tot he node list
		currentNodeList.AddRange(currentPath)
		'add this name of to the list
		currentNodeList.Add me.Name
		'add empty fields until the messageDepth
		dim i
		for i = currentNodeList.Count -1 to messageDepth -1  step +1
			currentNodeList.Add ""
		next
		'then add the other fields
		currentNodeList.Add me.Multiplicity
		'Add the name of the type 
		currentNodeList.Add me.TypeName
		'Add base type
		currentNodeList.Add me.BaseTypeName
		'add facets
		currentNodeList.Add getFacetsSpecification()
		'add the business attribute mapping and facets
		if me.CustomOrdering then
			if not me.IncludeDetails then
				'add empty fields for LDM mapping
				currentNodeList.Add "" 'Class
				currentNodeList.Add "" 'Attribute
				'add business attribute mapping details
				dim businessEntityNames
				businessEntityNames = ""
				dim businessAttributeNames
				businessAttributeNames = ""
				dim businessAttribute as EA.Attribute
				dim businessEntity as EA.Element
				for each businessAttribute in me.MappedBusinessAttributes
					'get the owner of the businessAttribute
					set businessEntity = Repository.GetElementByID(businessAttribute.ParentID)
					'set newline if needed
					if len(businessEntityNames) > 0 then
						businessEntityNames = businessEntityNames & vbNewLine
					end if
					'add businessEntityName
					businessEntityNames = businessEntityNames & businessEntity.Name
					'set the business attribute name
					if len(businessAttributeNames) > 0 then
						businessAttributeNames = businessAttributeNames & vbNewLine
					end if
					'add the businessAttributeName
					businessAttributeNames = businessAttributeNames & businessAttribute.Name
				next
				'add to the node list
				currentNodeList.Add businessEntityNames
				currentNodeList.Add businessAttributeNames
			end if
		'add the rules section
		elseif includeRules then
			if not validationRule is nothing then
				currentNodeList.Add validationRule.RuleId
				currentNodeList.Add validationRule.Name
				currentNodeList.Add validationRule.Reason
			else
				currentNodeList.Add ""
				currentNodeList.Add ""
				currentNodeList.Add ""
			end if
		end if
		'add the business usage section
		if me.CustomOrdering _
		and me.IncludeDetails _
		and not me.Message is nothing then
			dim fis as EA.Element
			for each fis in me.Message.Fisses 
				currentNodeList.Add ""
			next
		end if
		'return output
		set getThisNodeOutput = currentNodeList
	end function
	
	private function getFacetsSpecification()
		'initialize with empty string
		getFacetsSpecification = ""
		dim key
		for each key in me.Facets.Keys
			if len(getFacetsSpecification) > 0 then
				getFacetsSpecification = getFacetsSpecification & vbNewLine
			end if
			getFacetsSpecification = getFacetsSpecification & key & ": " & me.Facets.Item(key)
		next
		'for functional format the enum values should be included as well
		if not me.IncludeDetails  then
			dim enumType
			set enumType = getEnumType()
			if not enumType is nothing then
				dim enumValuesDescription
				enumValuesDescription = ""
				dim test as EA.Element
				dim enumValue as EA.Attribute
				for each enumValue in enumType.Attributes
					if len(enumValuesDescription) > 0 then
						'add newline
						enumValuesDescription = enumValuesDescription & vbNewLine
					end if
					'add the name
					enumValuesDescription = enumValuesDescription & enumValue.Name
					if me.CustomOrdering then
						'Description is stored in the tagged value CodeName
						dim tv as EA.AttributeTag
						for each tv in enumValue.TaggedValues
							if lcase(tv.Name) = "codename" then
								'add the description for this code
								enumValuesDescription = enumValuesDescription & " (" & tv.Value & ")"
								exit for
							end if
						next
					else
						'description is stored in the Alias
						enumValuesDescription = enumValuesDescription & " (" & enumValue.Alias & ")"
					end if
				next
				'add to facetSpecification
				if len(enumValuesDescription) > 0 then
					getFacetsSpecification = getFacetsSpecification & "Values allowed:" & VbNewLine & enumValuesDescription
				end if
			end if
		end if
	end function
	
	private function getEnumType()
		'initialize null
		set getEnumType = nothing
		'check if type element is enum
		if not me.TypeElement is nothing then
			if me.TypeElement.Type = "Enumeration" then
				set getEnumType = me.TypeElement
			end if
		end if
		if getEnumType is nothing _ 
		and not me.BaseTypeElement is nothing then
			if me.BaseTypeElement.Type = "Enumeration" then
				set getEnumType = me.BaseTypeElement
			end if
		end if
	end function
	
	'returns a list of all generalized elements of this elemnt
	private function getParents(childElement)
		dim directParents 
		dim sqlGetParents
		dim allParents
		set allParents = CreateObject("System.Collections.ArrayList")
		dim childElementID
		if not childElement is nothing then
			childElementID = childElement.ElementID
		else
			childElementID = me.ElementID
		end if
		sqlGetParents = "select c.End_Object_ID as Object_ID from t_connector c			 "  & _
						" where c.Connector_Type in ('Generalization','Generalisation')	 "  & _
						" and c.Start_Object_ID =" & childElementID
		set directParents = getElementsFromQuery(sqlGetParents)
		'add the direct parent to the list of all parents
		allParents.AddRange(directParents)
		'loop the parent and get their parents
		dim parent
		for each parent in directParents
			allParents.AddRange(getParents(parent))
		next
		'return
		set getParents = allParents
	end function
	'loads all Attribute notes both from this element as from its parents
	private function loadAllAttributeNodes(parents)
		'first load fro this element
		dim allAttributeNodes
		set allAttributeNodes = loadAttributeChildNodes(nothing)
		'then the one from the parents
		dim parent
		for each parent in parents
			allAttributeNodes.AddRange loadAttributeChildNodes(parent)
		next
	end function
	private function loadAttributeChildNodes(currentElement)
		set loadAttributeChildNodes = CreateObject("System.Collections.ArrayList")
		dim ownerElementID
		if not currentElement is nothing then
			ownerElementID = currentElement.ElementID
		else
			ownerElementID = me.ElementID
		end if
		'get attributes in the correct order (not for enum values
		dim SQLGetAttributes
		SQLGetAttributes = 	"select a.ID from (t_attribute a                             " & _
							" inner join t_object o on a.Object_ID = o.Object_ID)        " & _
							" where o.Object_Type <> 'Enumeration'                       " & _
							" and (o.Stereotype is null or o.Stereotype <> 'Enumeration')" & _
							" and a.Object_ID = " & ownerElementID & "                   " & _
							" and (a.Stereotype is null or a.Stereotype <> 'CON')        " & _
							" order by a.Pos, a.Name                                     "
		dim attributes
		set attributes = getattributesFromQuery(SQLGetAttributes)
		'loop the attributes
		dim attribute as EA.Attribute
		for each attribute in attributes
			'create the next messageNode
			dim newMessageNode
			set newMessageNode = new MessageNode
			newMessageNode.CustomOrdering = me.CustomOrdering
			newMessageNode.Message = me.Message
			'initialize
			newMessageNode.intitializeWithSource attribute, nothing, "", nothing, me
			'add to the childnodes list
			me.ChildNodes.Add newMessageNode
			'add to the output
			loadAttributeChildNodes.Add newMessageNode
		next
	end function
	
	'loads all Association nodes both from this element as from its parents
	private function loadAllAssociationNodes(parents)
		'first load for this element
		dim allAssociationNodes
		set allAssociationNodes = loadAssociationChildNodes(nothing)
		'then the ones from the parents
		dim parent
		for each parent in parents
			allAssociationNodes.AddRange loadAssociationChildNodes(parent)
		next
	end function
	
	private function loadAssociationChildNodes(currentElement)
		set loadAssociationChildNodes = CreateObject("System.Collections.ArrayList")
		dim ownerElementID
		if not currentElement is nothing then
			ownerElementID = currentElement.ElementID
		else
			ownerElementID = me.ElementID
		end if
		'get associations
		dim SQLAssociations
		SQLAssociations = 	"select c.Connector_ID from (t_connector c " & _
							" left join t_connectortag tv on (tv.ElementID = c.Connector_ID " & _
							" 						and tv.Property = 'sequencingKey')) " & _
							" where c.SourceIsAggregate > 0 " & _
							" and c.Start_Object_ID = " & ownerElementID & "  " & _         
							" order by tv.VALUE"
		dim associations
		set associations = getConnectorsFromQuery(SQLAssociations)
		'loop the associations
		dim association as EA.Connector
		for each association in associations
			'create the next messageNode
			dim newMessageNode
			set newMessageNode = new MessageNode
			newMessageNode.CustomOrdering = me.CustomOrdering
			newMessageNode.Message = me.Message
			'initialize
			newMessageNode.intitializeWithSource association.SupplierEnd, association, "", nothing, me
			'add to the childnodes list
			me.ChildNodes.Add newMessageNode
			'add to the output
			loadAssociationChildNodes.Add newMessageNode
		next
	end function
	
	private function setIsLeafNode()
		if not me.TypeElement is nothing then
			if me.TypeElement.Type = "Enumeration"_
			OR me.TypeElement.Stereotype = "Enumeration" _
			OR me.TypeElement.Stereotype = "XSDsimpleType" _
			OR me.TypeElement.Stereotype = "PRIM" then
				'enumerations and simple types are always leaf nodes
				m_IsLeafNode = true
			' a BDT is only a leafnode if it doesn't have any attributes except for the CON(tent)
			elseif me.TypeElement.Stereotype = "BDT" then
				m_IsLeafNode = true
				dim attribute as EA.Attribute
				for each attribute in me.TypeElement.Attributes
					if attribute.Stereotype <> "CON" then
						m_IsLeafNode = false
					end if
				next
			else
				m_IsLeafNode = false
			end if
		else
			m_IsLeafNode = true
		end if
	end function
	
end Class