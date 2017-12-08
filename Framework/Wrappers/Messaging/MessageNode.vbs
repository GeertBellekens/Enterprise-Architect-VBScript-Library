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
	End Sub
	
	'public properties
	
	' IsLeafNode property.
	Public Property Get IsLeafNode
		IsLeafNode = m_IsLeafNode
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
			if value <> me.TypeElement then
				me.TypeElement = nothing
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
				me.SourceElement = source
				me.Name = source.Name
				me.TypeElement = source
				me.Multiplicity = in_multiplicity
			case otAttribute
				me.SourceAttribute = source
				me.Name = source.Name
				dim attributeTypeObject
				set attributeTypeObject = nothing
				if source.ClassifierID > 0 then
					set attributeTypeObject = Repository.GetElementByID(source.ClassifierID)
					me.TypeElement = attributeTypeObject
				else
					me.TypeName = source.Type
				end if
			case otConnectorEnd
				me.SourceAssociationEnd = source
				if len(source.Role) > 0 then
					me.Name = source.Role
				else
					dim endObject as EA.Element
					'get the end object and use that name
					if source.End = "Supplier" then
						set endObject = Repository.GetElementByID(sourceConnector.SupplierID)
					else
						set endObject = Repository.GetElementByID(sourceConnector.ClientID)
					end if
					if not endObject is nothing then
						me.Name = endObject.Name
						me.TypeElement = endObject
					end if
				end if 
		end select
		'set the isLeafNode property
		setIsLeafNode
		'then load the child nodes
		if not me.IsLeafNode then
			loadChildNodes
		end if
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
		'load nested classes?
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
		'only add the name of the type if this is a leaf node
		if me.IsLeafNode then
			currentNodeList.Add me.TypeName
		else
			currentNodeList.Add ""
		end if
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
	'loads all Attribute notes both from this eleemnt as from its parents
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
							" order by a.Pos, a.Name                                     "
		dim attributes
		set attributes = getattributesFromQuery(SQLGetAttributes)
		'loop the attributes
		dim attribute as EA.Attribute
		for each attribute in attributes
			'create the next messageNode
			dim newMessageNode
			set newMessageNode = new MessageNode
			'initialize
			newMessageNode.intitializeWithSource attribute, nothing, "", nothing, me
			'add to the childnodes list
			me.ChildNodes.Add newMessageNode
			'add to the output
			loadAttributeChildNodes.Add newMessageNode
		next
	end function
	
	private function setIsLeafNode()
		if not me.TypeElement is nothing then
			if me.TypeElement.Type = "Enumeration"_
			OR me.TypeElement.Stereotype = "Enumeration" _
			OR me.TypeElement.Stereotype = "XSDsimpleType" then
				'enumerations and simple types are always leaf nodes
				m_IsLeafNode = true
			else
				m_IsLeafNode = false
			end if
		else
			m_IsLeafNode = true
		end if
	end function
	
end Class