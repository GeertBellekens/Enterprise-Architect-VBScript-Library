'[path=\Framework\Wrappers\Messaging]
'[group=Messaging]

!INC Utils.Include
!INC Local Scripts.EAConstants-VBScript
' Author: Geert Bellekens
' Purpose: A wrapper class for a message node in a messaging structure
' Date: 2017-03-14

const APIRemarkTagName = "remark API-Portal"
const DesignRemarkTagName = "FScustom"
const FSRemarkTagName = "FSremark"

const regularMessageContent = "regularMessageContent"
const functionalDesign = "functionalDesign"
const mappingDocument = "mappingDocument"


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
	Private m_SourcePackage
	Private m_ValidationRules
	Private m_IsLeafNode
	Private m_CustomOrdering
	Private m_Order
	Private m_Facets
	Private m_MappedBusinessAttributes
	Private m_IncludeDetails
	Private m_BaseTypeName
	Private m_BaseTypeElement
	Private m_Description
	Private m_APIRemark
	Private m_DesignRemark
	Private m_FSRemark
	Private m_ElementOwnerSourceType


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
		set m_SourcePackage = nothing
		set m_ValidationRules = CreateObject("System.Collections.ArrayList")
		m_IsLeafNode = false
		m_CustomOrdering = false
		m_order = 0
		set m_Facets = CreateObject("Scripting.Dictionary")
		set m_MappedBusinessAttributes = CreateObject("System.Collections.ArrayList")
		m_IncludeDetails = false
		m_BaseTypeName = ""
		set m_BaseTypeElement = nothing
		m_Description = ""
		set m_ElementOwnerSourceType = nothing
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
	
	Public Property Get ElementOwnerName
		if not me.ElementOwner is nothing then
			ElementOwnerName = me.ElementOwner.Name
		else
			ElementOwnerName = ""
		end if
	end Property
	
	Public Property Get ElementOwnerSourceType
		if m_ElementOwnerSourceType is nothing then
			if not me.ElementOwner is nothing then
				dim sourceGUID
				sourceGUID = getTaggedValueValue(me.ElementOwner, "sourceElement")
				if len(sourceGUID) > 0 then
					set m_ElementOwnerSourceType = Repository.GetElementByGuid(sourceGUID)
				end if
			end if
		end if
		set ElementOwnerSourceType = m_ElementOwnerSourceType
	end Property
	
	Public Property Get TypeElementSourceType
		dim sourceType
		set sourceType = nothing
		if not me.TypeElement is nothing then
			dim sourceGUID
			sourceGUID = getTaggedValueValue(me.TypeElement, "sourceElement")
			if len(sourceGUID) > 0 then
				set sourceType = Repository.GetElementByGuid(sourceGUID)
			end if
		end if
		set TypeElementSourceType = sourceType
	end Property
	
	Public property get ElementOwner
		dim m_elementOwner as EA.Element
		set m_elementOwner = nothing
		if not me.TypeElement is nothing then
			if lcase(me.TypeElement.Stereotype) = "xsdcomplextype" then
				set m_elementOwner = me.TypeElement
			end if
		end if
		if m_elementOwner is nothing _
		  and not me.ParentNode is nothing then
			set m_elementOwner = me.ParentNode.ElementOwner
		end if
		'return
		set ElementOwner = m_elementOwner
	end Property
	
	Public Property Get ComplexOrSimple
		if not me.ParentNode is nothing then
			ComplexOrSimple = "Simple"
		else
			ComplexOrSimple = "Simple"
		end if
		'if the type element is not complex type, then we return the parent name
		if not me.TypeElement is nothing then
			if lcase(me.TypeElement.Stereotype) = "xsdcomplextype" then
				ComplexOrSimple = "Complex"
			end if
		end if
	end Property
	
	public Property Get XSDName
		dim xsd
		xsd = ""
		'search source element package
		if not me.ElementOwner is nothing then
			xsd = getXSDName(me.ElementOwnerSourceType)
		elseif not me.ParentNode is nothing then
			xsd = me.ParentNode.XSDName
		end if
		'if still empty then we return the xsdName of the typeElement
		if len(xsd) = 0 _
		  and not me.TypeElement is nothing then
			xsd = getXSDName(me.TypeElementSourceType)
		end if
		'return
		XSDName = xsd
	end Property
	
	function getXSDName(sourceType)
		dim xsd
		xsd = ""
		if not sourceType is nothing then
			dim sourcePackage as EA.Package
			set sourcePackage = Repository.GetPackageByID(sourceType.PackageID)
			xsd = sourcePackage.Name
			'check if we need to add suffix .Types or .BaseTypes
			if lcase(right(sourcePackage.Name, len(".BaseTypes"))) <> ".basetypes" then
				dim parentSourcePackage as EA.Package
				set parentSourcePackage = Repository.GetPackageByID(sourcePackage.ParentID)
				if lcase(parentSourcePackage.Name) = "basetypes" then
					xsd = xsd & ".BaseTypes"
				else
					xsd = xsd & ".Types"
				end if
			end if
		end if
		'return
		getXSDName = xsd
	end function
	
	Public Property Get Nillable
		dim m_Nillable
		m_Nillable = false
		if not me.sourceAttribute is nothing then
			if lcase (getTaggedValueValue(me.sourceAttribute, "nillable")) = "true" then
				m_Nillable = true
			end if
		end if
		Nillable = m_Nillable
	end Property
	
	Public Property Get DefaultValue
		dim m_DefaultValue
		m_DefaultValue = ""
		if not me.sourceAttribute is nothing then
			m_DefaultValue = getTaggedValueValue(me.sourceAttribute, "default")
			if lcase(m_defaultValue) = "true" _
			  or lcase(m_defaultValue) = "false" then
			  'avoid translation of true or false in excel (WAAR/ONWAAR) by adding a single quote
			  m_defaultValue = "'" & m_defaultValue
			end if
		end if
		DefaultValue = m_DefaultValue
	end Property
	
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
			if len(sourceAssociationEnd.Cardinality) = 1 then
				if sourceAssociationEnd.Cardinality = "*" then
					returnedMultiplicity = "0..*"
				else
					returnedMultiplicity = sourceAssociationEnd.Cardinality & ".." & sourceAssociationEnd.Cardinality
				end if
			elseif len(sourceAssociationEnd.Cardinality) = 0 then
				'default multiplicity = 1..1
				returnedMultiplicity = "1..1"
			else
				returnedMultiplicity = sourceAssociationEnd.Cardinality
			end if
		end if
		'return the actual value
		Multiplicity = returnedMultiplicity
	End Property
	Public Property Let Multiplicity(value)
		if not me.SourceElement is nothing then
			m_Multiplicity = value
		end if
	End Property
	
	Public Property Get LowerBound
		LowerBound = left(me.Multiplicity, 1)
	end Property
	
	Public Property Get UpperBound
		UpperBound = right(me.Multiplicity, 1)
	end Property
	
	private function determineMultiplicity(lower,upper,defaultLower, defaultUpper)
		'check to make sur the values are filled in and replace them with the default values if not the case
		if len(lower) = 0 then
			lower = defaultLower
		end if
		if len(upper) = 0 then
			upper = defaultUpper
		elseif upper = "-1" then
			upper = "*"
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
	
	' SourcePackage property.
	Public Property Get SourcePackage
		set SourcePackage = m_SourcePackage
	End Property
	Public Property Let SourcePackage(value)
		set m_SourcePackage = value
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
	
	' APIRemark property.
	Public Property Get APIRemark
		if IsEmpty(m_APIRemark) then
			'not initialized, get value
			m_APIRemark = getRemark(me.Source, APIRemarkTagName )
			'if remark info not found then look a the tagged value from the basetype elemnt
			if len(m_APIRemark) = 0 then
				m_APIRemark = getRemark(me.TypeElement, APIRemarkTagName)
			end if
		end if
		APIRemark = m_APIRemark
	End Property
	
	' DesignRemark property.
	Public Property Get DesignRemark
		if IsEmpty(m_DesignRemark) then
			'not initialized, get value
			'check if source has steroeotype Redefine%
			m_DesignRemark = getRedefineStereotype(me.Source)
			'check if the base type has a redefine stereotype
			dim typeStereo
			typeStereo = getRedefineStereotype(me.TypeElement)
			if len(typeStereo) > 10 and len(m_DesignRemark) > 0 then
				m_DesignRemark = m_DesignRemark & ", "
			end if
			m_DesignRemark = m_DesignRemark & typeStereo
		end if
		DesignRemark = m_DesignRemark
	End Property
	
	private function getRedefineStereotype(item)
		dim redefineStereo
		redefineStereo = ""
		if item is nothing then
			getRedefineStereotype = redefineStereo
			exit function
		end if
		dim stereotypes
		stereotypes = Split(item.StereotypeEx, ",")
		dim stereotype
		for each stereotype in stereotypes
			if instr(lcase(stereotype), "redefine") > 0 then
				redefineStereo = stereotype
				exit for 'we only expect one
			end if
		next
		'return
		getRedefineStereotype = redefineStereo
	end function
	
	' FSRemark property.
	Public Property Get FSRemark
		if IsEmpty(m_FSRemark) then
			'not initialized, get value
			m_FSRemark = getRemark(me.Source, FSRemarkTagName )
			'if remark info not found then look a the tagged value from the basetype elemnt
			if len(m_FSRemark) = 0 then
				m_FSRemark = getRemark(me.TypeElement, FSRemarkTagName)
			end if
		end if
		FSRemark = m_FSRemark
	End Property
	
	Public Property Get Source
		'Extra functional information
		if not me.SourceAttribute is nothing then
			set Source = me.SourceAttribute
		elseif not me.SourceAssociationEnd is nothing then
			set Source = me.SourceAssociationEnd
		elseif not SourceElement is nothing then
			set Source = me.SourceElement
		else 
			set Source = nothing
		end if
	End Property
	
	' MappedBusinessAttributes property.
	Public Property Get MappedBusinessAttributes
		set MappedBusinessAttributes = m_MappedBusinessAttributes
	End Property
	Public Property Let MappedBusinessAttributes(value)
		set m_MappedBusinessAttributes = value
	End Property
	
	' Description property.
	Public Property Get Description
		Description = m_Description
	End Property
	Public Property Let Description(value)
		m_Description = Repository.GetFormatFromField("TXT", value)
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
			case otPackage
				setPackageNode source
		end select
		Repository.WriteOutput outPutName, now() & " Processing node '" & me.Name & "'", 0
		'set the isLeafNode property
		setIsLeafNode
		'then load the child nodes
		if not me.IsLeafNode then
			loadChildNodes
		end if
	end function
	
	private function setPackageNode(source)
		me.SourcePackage = source
		me.Name = source.Name
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
		'set the name
		me.Name = source.Name
		'set the description
		me.Description = source.Notes
		'set the order
		me.Order = getSequencingKey(source)
		'set the type
		dim attributeTypeObject as EA.Element
		set attributeTypeObject = nothing
		if source.ClassifierID > 0 then
			on error resume next
			set attributeTypeObject = Repository.GetElementByID(source.ClassifierID)
			if Err.Number <> 0 then
				'classifier not found
				Err.Clear
			end if
			on error goto 0
		end if
		if not attributeTypeObject is nothing then
			'regular attribute
			me.TypeElement = attributeTypeObject
			'find the parent class (not for enumerations)
			if not (me.TypeElement.Type = "Enumeration" _
				or lcase(me.TypeElement.Stereotype) = "enumeration") then
				dim parentClass as EA.Element
				if attributeTypeObject.BaseClasses.Count > 0 then
					for each parentClass in attributeTypeObject.BaseClasses
						me.BaseTypeElement = parentClass
						exit for 'return immediate
					next
				else
					'get the name of the parent from the genlinks field
					dim parentName
					dim genLinks
					if len(attributeTypeObject.Genlinks) > 0 then
						genLinks = attributeTypeObject.Genlinks
						parentName = getValueForkey(genLinks, "Parent")
						if len(parentName) > 0 then
							me.BaseTypeName = parentName
						end if
					end if
				end if
			end if
		else
			me.TypeName = source.Type
		end if
		'first get the facets from the type element
		if not me.typeElement is nothing then
			getFacets typeElement
		end if
		'get the facets
		getFacets source
	end function
		
	'gets the facets from an attribute
	private function getFacets(sourceAttribute)
		dim t as EA.Attribute t.TaggedValues.de
		dim tv as EA.TaggedValue
		'first loop the standard facets
		for each tv in sourceAttribute.TaggedValues
			if len(tv.Value) > 0 and _
			  (tv.Name = "enumeration" _
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
		me.Order = getSequencingKey(sourceConnector.ClientEnd)
'		if me.Order = 0 then 
'			'association go in the back if there is no explicit order defined
'			me.Order = 999
'		end if
		dim endObject as EA.Element
		'get the end object 
		if source.End = "Supplier" then
			set endObject = Repository.GetElementByID(sourceConnector.SupplierID)
		else
			set endObject = Repository.GetElementByID(sourceConnector.ClientID)
		end if
		'set the name = name of role + name of end object + remove underscores
		if len(source.Role) > 0 then
			me.Name = source.Role
		else
			'use the end object name as rolename
			me.Name = endObject.Name
		end if 
		'set the type
		me.TypeElement = endObject
		'set the description
		if len(sourceConnector.Notes) > 0 then
			me.Description = sourceConnector.Notes
		else
			me.Description = endObject.Notes
		end if
	end function
	
	public function getSequencingKey(sourceItem)
		'initialize at 999
		getSequencingKey = 999
		dim taggedValue as EA.TaggedValue
		for each taggedValue in sourceItem.TaggedValues
			dim tagName
			if sourceItem.ObjectType = otConnectorEnd then
				tagName = taggedValue.Tag
			else
				tagName = taggedValue.Name
			end if
			if Lcase(tagName) = "position" then
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
		if not me.SourcePackage is nothing then
			'get all root elements as childnodes, and call LoadChildNodes on all of them
			loadPackageChildNodes
		else
			'first remember the list of parent elements
			dim parents
			set parents = getParents(nothing)
			'TODO: load in correct order?
			'load attributes
			dim attributeNodes
			set attributeNodes = loadAllAttributeNodes(parents)
			'load associations
			dim associationNodes
			set associationNodes = loadAllAssociationNodes(parents)
			'TODO: load nested classes?
			resetOrders attributeNodes, associationNodes
			'reorder nodes
			reOrderNodes me.ChildNodes
		end if
	end function
	
	private function loadPackageChildNodes
		'get all root elements as childnodes, and call LoadChildNodes on all of them
		dim packageIDtree
		packageIDtree = getPackageTreeIDString(me.SourcePackage)
		dim rootElements
		dim sqlGetData
		sqlGetData = "select o.Object_ID                                                            " & vbNewLine & _
					" from t_object o                                                              " & vbNewLine & _
					" where  o.Package_ID in (" & packageIDtree & ")                               " & vbNewLine & _
					" and o.Stereotype like 'XSDComplexType'                                       " & vbNewLine & _
					" and not exists (select t.Object_ID from t_object t                           " & vbNewLine & _
					" 				where t.Package_ID in (" & packageIDtree & ")                  " & vbNewLine & _
					" 				and t.Stereotype = 'XSDtopLevelElement'                        " & vbNewLine & _
					" 				union                                                          " & vbNewLine & _
					" 				select t.Object_ID from t_object t                             " & vbNewLine & _
					" 				inner join t_attribute a on a.Object_ID = t.Object_ID          " & vbNewLine & _
					" 									and a.Classifier = o.Object_ID             " & vbNewLine & _
					" 				where t.Package_ID in (" & packageIDtree & ")                  " & vbNewLine & _
					" 				union                                                          " & vbNewLine & _
					" 				select t.Object_ID from t_object t                             " & vbNewLine & _
					" 				inner join t_connector c on c.Start_Object_ID = t.Object_ID    " & vbNewLine & _
					" 										and c.End_Object_ID = o.Object_ID      " & vbNewLine & _
					" 										and c.Connector_Type = 'Association'   " & vbNewLine & _
					" 				where t.Package_ID in (" & packageIDtree & ")                  " & vbNewLine & _
					" 				)                                                              "
		set rootElements = getElementsFromQuery(sqlGetData)
		dim rootElement as EA.Element
		for each rootElement in rootElements
			'create the next messageNode
			dim newMessageNode
			set newMessageNode = new MessageNode
			newMessageNode.CustomOrdering = me.CustomOrdering
			'initialize
			newMessageNode.intitializeWithSource rootElement, nothing, "1..1", nothing, me
			'add to the childnodes list
			me.ChildNodes.Add newMessageNode
		next
	end function
	
	private function resetOrders(attributeNodes, associationNodes)
		'the first ones are the ones that have their position set.
		'after that we put the associations to XSDChoice elements (alphabetically) and then the rest alphabetically
		reOrderNodes me.ChildNodes
		'get the maximum position in all childNodes
		dim maxPosition
		maxPosition = getMaxPosition()
		'get all XSDChoiceNodes
		dim xsdChoiceNodes
		set xsdChoiceNodes = getXSDChoiceNodes()
		'set the orders in the XSDChoiceNodes starting at the max position
		setOrders xsdChoiceNodes, maxPosition
		'reorder the childnodes again
		reOrderNodes me.ChildNodes

	end function
	
	function printDebugInfo(nodes, Message)
		Session.Output Message
		dim node
		for each node in nodes
			Session.Output "Node Name: '" & node.Name & "' Node Order: " & node.Order
		next
	end function
	
	function setOrders(nodes, startOrder)
		dim node
		dim currentOrder
		currentOrder = startOrder
		for each node in nodes
			currentOrder = currentOrder + 1
			node.Order = currentOrder
		next
	end function
	
	function getXSDChoiceNodes()
		dim xsdChoiceNodes
		set xsdChoiceNodes = CreateObject("System.Collections.ArrayList")
		dim childNode
		'get all elements whose type element has steroetype XSDChoice
		for each childNode in me.ChildNodes
			if not childNode.TypeElement is nothing _
			 And childNode.Order = 999 then
				if childNode.TypeElement.HasStereotype("XSDchoice") then
					xsdChoiceNodes.Add childNode
				end if
			end if
		next
		'reorder the XSDChoiceNodes
		reOrderNodes xsdChoiceNodes
		'return
		set getXSDChoiceNodes = xsdChoiceNodes
	end function
	
	function getMaxPosition()
		dim maxPosition
		maxPosition = 0
		dim childNode 
		for each childNode in me.ChildNodes
			'stop if we have a node with order 999
			if childNode.Order = 999 then
				exit for
			elseif childNode.Order > maxPosition then
				maxPosition = childNode.Order 
			end if
		next
		'return
		getMaxPosition = maxPosition
	end function
	
		
	public function reOrderNodes(nodesList)
		dim childNode
			dim i
		dim goAgain
		goAgain = false
		dim currentNode
		dim nextNode
		for i = 0 to nodesList.Count -2 step 1
			set currentNode = nodesList.Item(i)
			set nextNode = nodesList.Item(i +1)
			if  currentNode.Order > nextNode.Order or _
				(currentNode.Order = nextNode.Order and currentNode.Name > nextNode.Name)  then
				nodesList.RemoveAt(i +1)
				nodesList.Insert i, nextNode
				goAgain = true
			end if
		next
		'if we had to swap an element then we go over the list again
		if goAgain then
			reOrderNodes nodesList
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
	
	public function getStructureOutput(current_order,currentPath,messageDepth, exportType)
		'create the output
		dim nodeOutputList
		set nodeOutputList = CreateObject("System.Collections.ArrayList")
		dim currentNodeList
		select case exportType
			case functionalDesign
				set currentNodeList = getThisNodeFSOutput(current_order,currentPath, messageDepth)
			case mappingDocument
				set currentNodeList = getThisNodeMappingOutput(current_order,currentPath, messageDepth)
			case else
				set currentNodeList = getThisNodeStructureOutput(current_order,currentPath, messageDepth)
		end select
		'up or the order number
		current_order = current_order + 1
		'add the list for this node to the output
		nodeOutputList.Add currentNodeList
		'add this node to the currentPath
		dim mycurrentpath
		set myCurrentPath = CreateObject("System.Collections.ArrayList")
		myCurrentPath.AddRange(currentPath)
		myCurrentPath.Add me.Name
		'get the output for the child nodes
		dim childNode
		for each childNode in me.ChildNodes
			if exportType = functionalDesign OR _
			  exportType = mappingDocument OR _
			  childNode.ChildNodes.Count > 0 then
				dim childOutPut
				set childOutPut = childNode.getStructureOutput(current_order,myCurrentPath,messageDepth, exportType)
				nodeOutputList.AddRange(childOutPut)
			end if
		next
		'return list
		set getStructureOutput = nodeOutputList
	end function
	'gets the output format for this node and its childnodes
	public function getOutput(current_order,currentPath,messageDepth, includeRules)
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
			set childOutPut = childNode.getOutput(current_order,myCurrentPath,messageDepth, includeRules)
			nodeOutputList.AddRange(childOutPut)
		next
		'return list
		set getOutput = nodeOutputList
	end function
	
	private function getThisNodeStructureOutput(current_order,currentPath, messageDepth)
		'get the list for this node
		dim currentNodeList
		set currentNodeList = CreateObject("System.Collections.ArrayList")
		'add this name of to the list
		currentNodeList.Add space(currentPath.Count * 6 ) & me.Name
		'add the multiplicity
		currentNodeList.Add "[" & me.Multiplicity & "]"
		'return output
		set getThisNodeStructureOutput = currentNodeList
	end function
	
	
	function getThisNodeMappingOutput(current_order,currentPath, messageDepth)
		'get the list for this node
		dim currentNodeList
		set currentNodeList = CreateObject("System.Collections.ArrayList")
		'add the full path to the list
		if currentPath.Count > 0 then
			currentNodeList.Add Join(currentPath.ToArray(), ".") & "." & me.Name
		else
			currentNodeList.Add me.Name
		end if
		'lowerbound
		dim minOccursString
		minOccursString = me.LowerBound
		if me.Nillable then
			minOccursString = minOccursString & " (nillable)"
		end if
		currentNodeList.Add minOccursString
		'upperBound
		dim maxOccursString
		if me.UpperBound = "*" then
			maxOccursString = "unbounded"
		else
			maxOccursString = me.UpperBound
		end if
		currentNodeList.Add maxOccursString
		'original XSD namespace
		currentNodeList.Add me.xsdName
		'element type name
		currentNodeList.Add me.ElementOwnerName
		'type
		if me.ComplexOrSimple = "Complex" then
			currentNodeList.Add "Complex"
		elseif lcase(right(me.TypeName, len("choice"))) = "choice" then
			currentNodeList.Add "Choice"
		else
			currentNodeList.Add me.TypeName
		end if
		'default
		currentNodeList.Add me.DefaultValue
		'annotation
		currentNodeList.Add me.Description
		'Design remark
		currentNodeList.Add me.DesignRemark
		'add the API remark
		currentnodeList.Add me.APIRemark
		'return output
		set getThisNodeMappingOutput = currentNodeList
	end function
	
	function getThisNodeFSOutput(current_order,currentPath, messageDepth)
		'get the list for this node
		dim currentNodeList
		set currentNodeList = CreateObject("System.Collections.ArrayList")
		'add this name of to the list
		currentNodeList.Add space(currentPath.Count * 5 ) & me.Name
		'lowerbound
		dim minOccursString
		if me.ParentNode is nothing then
			'no minoccurs for root nodes
			minOccursString = ""
		else
			minOccursString = me.LowerBound
			if me.Nillable then
				minOccursString = minOccursString & " (nillable)"
			end if
		end if
		currentNodeList.Add minOccursString
		'upperBound
		dim maxOccursString
		if me.ParentNode is nothing then
			'no maxOccurs for root nodes
			maxOccursString = ""
		else
			if me.UpperBound = "*" then
				maxOccursString = "unbounded"
			else
				maxOccursString = me.UpperBound
			end if
		end if
		currentNodeList.Add maxOccursString
		'original XSD namespace
		currentNodeList.Add me.xsdName
		'element original type name
		if not me.ElementOwnerSourceType is nothing then
			currentNodeList.Add me.ElementOwnerSourceType.Name
		else
			currentNodeList.Add ""
		end if
		'type
		if me.ParentNode is nothing then
			'no type for root nodes
			currentNodeList.Add ""
		else
			if me.ComplexOrSimple = "Complex" then
				currentNodeList.Add "Complex"
			elseif lcase(right(me.TypeName, len("choice"))) = "choice" then
				currentNodeList.Add "Choice"
			else
				currentNodeList.Add me.TypeName
			end if
		end if
		'default
		currentNodeList.Add me.DefaultValue
		'annotation
		currentNodeList.Add me.Description
		'Design remark
		currentNodeList.Add me.DesignRemark
		'Design remark
		currentNodeList.Add me.FSRemark
		'add the API remark
		currentnodeList.Add me.APIRemark
		'return output
		set getThisNodeFSOutput = currentNodeList
	end function
	
	private function getThisNodeOutput(current_order,currentPath, messageDepth,validationRule, includeRules)
		'get the list for this node
		dim currentNodeList
		set currentNodeList = CreateObject("System.Collections.ArrayList")
		'add this name of to the list
		currentNodeList.Add space(currentPath.Count * 6 ) & me.Name
		'add the multiplicity
		currentNodeList.Add "[" & me.Multiplicity & "]"
		'add the description
		currentnodeList.Add me.Description
		'The type field should contain:
		' - Nothing if the type element is a complex type
		' - The typeName if the type element is one of the standard xsd base type
		' - The base type name + facets if the type element is a simple type
		' - TypeName if TypeElement is nothing
		dim outputTypeName
		outputTypeName = ""
		if not me.typeElement is nothing then
			if lcase(me.typeElement.Stereotype) = "xsdsimpletype" then
				'get the parent package
				dim parentPackage
				set parentPackage = Repository.GetPackageByID(me.typeElement.PackageID)
				if lcase(parentPackage.Name) = "xsddatatypes" then
					outputTypeName = me.TypeName
				else
					outputTypeName = me.BaseTypeName
				end if
				'add the facets
				outputTypeName = outputTypeName & getFacetsSpecification()
			elseif me.typeElement.Type = "Enumeration" then
				outputTypeName = me.TypeName
			end if
		else
			outputTypeName = me.TypeName
		end if
		'add it to the node list
		currentnodeList.Add outputTypeName
		'add the functional info to the node list
		currentnodeList.Add me.APIRemark
		'add the rules section
		if includeRules then
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
		'return output
		set getThisNodeOutput = currentNodeList
	end function
	
	function getRemark(element, tagName)
		getRemark = "" 'initial value
		if not element is nothing then
			dim tv as EA.TaggedValue
			for each tv in element.TaggedValues
				if lcase(tv.Name) = lcase(tagName) then '"remark API-Portal"
					getRemark = tv.Value
					exit for
				end if
			next
		end if
	end function
	
	private function getFacetsSpecification()
		'initialize with empty string
		getFacetsSpecification = ""
		dim key
		for each key in me.Facets.Keys
			if len(getFacetsSpecification) > 0 then
				getFacetsSpecification = getFacetsSpecification & vbNewLine
			else 
				getFacetsSpecification = " with "
			end if
			getFacetsSpecification = getFacetsSpecification & key & ": " & me.Facets.Item(key)
		next
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
		set loadAllAttributeNodes = allAttributeNodes
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
		SQLGetAttributes = 	"select a.ID                                                  " & _
							" from t_attribute a                                          " & _
							" inner join t_object o on a.Object_ID = o.Object_ID          " & _
							" left join t_attributetag atv on atv.ElementID = a.ID        " & _
							" 								and atv.Property = 'position' " & _
							" 								and isnumeric(atv.VALUE) = 1  " & _
							" where o.Object_Type <> 'Enumeration'                        " & _
							" and (o.Stereotype is null or o.Stereotype <> 'Enumeration') " & _
							" and a.Object_ID = " & ownerElementID & "                    " & _
							" order by CONVERT(int,isnull(atv.value,999)), a.Pos, a.Name  "
		dim attributes
		set attributes = getattributesFromQuery(SQLGetAttributes)
		'loop the attributes
		dim attribute as EA.Attribute
		for each attribute in attributes
			'create the next messageNode
			dim newMessageNode
			set newMessageNode = new MessageNode
			newMessageNode.CustomOrdering = me.CustomOrdering
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
		'return
		set loadAllAssociationNodes = allAssociationNodes
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
		SQLAssociations = 	"select c.Connector_ID from t_connector c " & _
							" left join t_taggedvalue tv on tv.ElementID = c.ea_guid            " & _
							" 							and tv.BaseClass = 'ASSOCIATION_SOURCE' " & _
							" 							and tv.TagValue = 'position'            " & _
							" where c.Connector_Type = 'Association' " & _
							" and c.Start_Object_ID = " & ownerElementID & "  " & _         
							" order by cast(isnull(tv.Notes,999) as int)"
		dim associations
		set associations = getConnectorsFromQuery(SQLAssociations)
		'loop the associations
		dim association as EA.Connector
		for each association in associations
			'create the next messageNode
			dim newMessageNode
			set newMessageNode = new MessageNode
			newMessageNode.CustomOrdering = me.CustomOrdering
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
		elseif not me.sourcePackage is nothing then
			m_IsLeafNode = false
		else
			m_IsLeafNode = true
		end if
	end function
	
end Class