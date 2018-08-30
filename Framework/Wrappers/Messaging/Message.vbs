'[path=\Framework\Wrappers\Messaging]
'[group=Messaging]

!INC Utils.Include

' Author: Geert Bellekens
' Purpose: A wrapper class for a message node in a messaging structure
' Date: 2017-03-14

Class Message
	'private variables
	Private m_Name
	Private m_RootNode
	Private m_MessageDepth
	Private m_BaseTypes
	Private m_Enumerations
	Private m_Prefix
	private m_ValidationRules
	Private m_CustomOrdering
	Private m_IncludeDetails
	Private m_Fisses
	
	'constructor
	Private Sub Class_Initialize
		m_Name = ""
		set m_RootNode = nothing
		m_MessageDepth = 0
		set m_BaseTypes = CreateObject("Scripting.Dictionary")
		set m_Enumerations = CreateObject("Scripting.Dictionary")
		m_Prefix = ""
		set m_ValidationRules = CreateObject("System.Collections.ArrayList")
		m_CustomOrdering = false
		m_IncludeDetails = false
		set m_Fisses = nothing
	End Sub
	
	'public properties
	
	' Name property.
	Public Property Get Name
	  Name = m_Name
	End Property
	Public Property Let Name(value)
	  m_Name = value
	End Property
	
	' RootNode property.
	Public Property Get RootNode
	  set RootNode = m_RootNode
	End Property
	Public Property Let RootNode(value)
	  set m_RootNode = value
	End Property
	
	' MessageDepth property.
	Public Property Get MessageDepth
		if m_MessageDepth = 0 then
			m_MessageDepth = getMessageDepth()
		end if
		MessageDepth = m_MessageDepth
	End Property
	
	' BaseTypes property.
	Public Property Get BaseTypes
	  set BaseTypes = m_BaseTypes
	End Property
	
	' Enumerations property.
	Public Property Get Enumerations
	  set Enumerations = m_Enumerations
	End Property
	
	' Prefix property.
	Public Property Get Prefix
	  Prefix = m_Prefix
	End Property
	
	' ValidationRules property.
	Public Property Get ValidationRules
	  set ValidationRules = m_ValidationRules
	End Property
	Public Property Let ValidationRules(value)
	  set m_ValidationRules = value
	End Property
	
	'CustomOrdering
	Public Property Get CustomOrdering
	  CustomOrdering = m_CustomOrdering
	End Property
	Public Property Let CustomOrdering(value)
		m_CustomOrdering = value
	End Property
	
	' IncludeDetails property.
	Public Property Get IncludeDetails
		IncludeDetails = m_IncludeDetails
	End Property
	Public Property Let IncludeDetails(value)
		m_IncludeDetails = value
	End Property
	
	' Fisses property.
	Public Property Get Fisses
		if m_Fisses is nothing then
			loadFisses
		end if
		set Fisses = m_Fisses
	End Property
	Public Property Let Fisses(value)
		set Fisses = value
	End Property
	
	private function loadFisses()
		set m_Fisses = CreateObject("System.Collections.ArrayList")
		'MA(subset) -trace-> MA(messaging Model) -realize-> Message -realize-> FIS
		dim getFissesSQL
		getFissesSQL = "select fis.Object_ID from t_object o                                             " & _
					" inner join t_connector c on c.Start_Object_ID = o.Object_ID                      " & _
					" 							and c.Connector_Type = 'Abstraction'                   " & _
					" 							and c.Stereotype = 'trace'                             " & _
					" inner join t_object om on om.Object_ID = c.End_Object_ID                         " & _
					" 						and om.Name = o.Name                                       " & _
					" 						and om.Object_Type = o.Object_Type                         " & _
					" inner join t_connector omc on omc.Start_Object_ID = om.Object_ID                 " & _
					" 						 and omc.Connector_Type in ('Realization', 'Realisation')  " & _
					" inner join t_object msg on msg.object_ID = omc.End_Object_ID                     " & _
					" 						and msg.Object_Type = 'Class'                              " & _
					" 						and msg.Stereotype = 'Message'                             " & _
					" inner join t_connector msgc on msgc.Start_Object_ID = msg.Object_ID              " & _
					" 						 and msgc.Connector_Type in ('Realization', 'Realisation') " & _
					" inner join t_object fis on fis.Object_ID = msgc.End_Object_ID                    " & _
					" 						and fis.Object_Type = 'Class'                              " & _
					" 						and fis.Stereotype = 'Message'                             " & _
					" where o.Object_ID = " & me.RootNode.ElementID
		dim fisses
		set fisses = getElementsFromQuery(getFissesSQL)
		dim fis as EA.Element
		for each fis in fisses
			m_Fisses.Add fis
		next
	end function
	
	public function loadMessage(eaRootNodeElement)
		
		'set the name of the message
		'the name of the message is equal to the name of the owning package
		dim ownerPackage as EA.Package
		set ownerPackage = Repository.GetPackageByID(eaRootNodeElement.PackageID)
		me.Name = ownerPackage.Name
		'set the prefix
		m_Prefix = getPrefix(ownerPackage)
		'set the customOrdering property (check if «MA» is one of the stereotypes
		dim rootNodeStereotypes
		dim rootNodeStereotype
		rootNodeStereotypes = split(eaRootNodeElement.StereotypeEx, ",")
		for each rootNodeStereotype in rootNodeStereotypes
			if rootNodeStereotype = "MA" then
				me.CustomOrdering = true
				'for message assemblies the name is stored on the element
				me.Name = eaRootNodeElement.Name
				exit for
			end if
		next
		'create the root node
		me.RootNode = new MessageNode
		me.RootNode.CustomOrdering = me.CustomOrdering
		me.RootNode.IncludeDetails = me.IncludeDetails
		me.RootNode.Message = me
		me.RootNode.intitializeWithSource eaRootNodeElement, nothing, "1..1", nothing, nothing
		setBaseTypesAndEnumerations(me.RootNode)
		'link the message validation rules
		getMessageValidationRules()
	end function
	
	private function getPrefix(ownerPackage)
		getPrefix = ""
		dim taggedValue as EA.TaggedValue
		for each taggedValue in ownerPackage.Element.TaggedValues
			if taggedValue.Name = "targetNamespacePrefix" then
				getPrefix = taggedValue.Value
				exit for
			end if
		next
	end function
	
	private function getMessageValidationRules()
		dim getRulesElementsSQL
		getRulesElementsSQL = 	"select r.* from ((t_object o                                     " & _
								" inner join t_connector c on (c.End_Object_ID = o.Object_ID      " & _
								" 							and c.Connector_Type = 'Dependency' ))" & _
								" inner join t_object r on (c.Start_Object_ID = r.Object_ID       " & _
								" 						and r.Object_Type = 'Test'                " & _
								" 						and r.Stereotype = 'Message Test Rule'))  " & _
								" where o.Object_ID = " & me.RootNode.ElementID
		dim rulesElements
		set rulesElements = getElementsFromQuery(getRulesElementsSQL)
		dim rulesElement
		for each rulesElement in rulesElements
			dim validationRule
			set validationRule = new MessageValidationRule
			validationRule.initialiseWithTestElement(rulesElement)
			m_ValidationRules.Add validationRule
			'find the node this rule applies to ad add it to that node
			me.RootNode.linkRuletoNode validationRule, validationRule.Path
		next
	end function
	
	
	private function setBaseTypesAndEnumerations(messageNode)
		'check if messageNode is leafNode
		if messageNode.IsLeafNode then
			dim foundEnumeration
			foundEnumeration = false
			'check if the typeElement is an enumeration
			if not messageNode.TypeElement is nothing then
				if messageNode.TypeElement.Type = "Enumeration"_
				OR messageNode.TypeElement.Stereotype = "Enumeration" then
					foundEnumeration = true
					if not me.Enumerations.Exists(messageNode.TypeName) then
						'add to enumerations list
						me.Enumerations.Add messageNode.TypeName, messageNode.TypeElement
					end if
				end if
			end if
			'if we haven't found an enumeration we add the type to the basetypes
			if not foundEnumeration then
				if not me.BaseTypes.Exists(messageNode.TypeName) then
					'add to BaseTypes list
					me.BaseTypes.Add messageNode.TypeName, messageNode.TypeElement
				end if
			end if
		else
			'not a leafnode, check the childnodes
			dim childNode
			for each childNode in messageNode.ChildNodes
				setBaseTypesAndEnumerations childNode 
			next
		end if
		'add the base type to the list of types
		if not messageNode.BaseTypeElement is nothing then
			if messageNode.BaseTypeElement.Type = "Enumeration"_
			OR messageNode.BaseTypeElement.Stereotype = "Enumeration"  then
				if not me.Enumerations.Exists(messageNode.BaseTypeName) then
					'add to enumerations list
					me.Enumerations.Add messageNode.BaseTypeName, messageNode.BaseTypeElement
				end if
			else
				'add as base type
				if not me.BaseTypes.Exists(messageNode.BaseTypeName) then
					me.BaseTypes.Add messageNode.BaseTypeName, messageNode.BaseTypeElement
				end if
			end if
		end if
	end function
	
	'create an arraylist of arraylists with the details of this message
	public function createOuput(includeRules)
		dim outputList
		'create empty list for current path
		dim currentPath
		set currentPath = CreateObject("System.Collections.ArrayList")
		'start with the rootnode
		set outputList = me.RootNode.getOuput(1,currentPath,me.MessageDepth, includeRules)
		'return outputlist
		set createOuput = outputList
	end function
	
	'create an arraylist of arraylists with the details of this message
	public function createUnifiedOutput(includeRules, depth)
		dim outputList
		'create empty list for current path
		dim currentPath
		set currentPath = CreateObject("System.Collections.ArrayList")
		'start with the rootnode
		set outputList = me.RootNode.getOuput(1,currentPath,depth, includeRules)
		'return outputlist
		set createUnifiedOutput = outputList
	end function
	
	'create an arraylist of arraylists with the details of this message including he headers
	public function createFullOutput(includeRules)
		dim fullOutput
		dim headers
		set fullOutput = me.createOuput(includeRules)
		set headers = getHeaders(includeRules)
		'insert the headers before the rest of the output
		fullOutput.Insert 0, headers
		set createFullOutput = fullOutput
	end function
	
	'gets the maximum depth of this message
	private function getMessageDepth()
		dim message_depth
		message_depth = 0
		message_depth = me.RootNode.getDepth(message_depth)
		getMessageDepth = message_depth
	end function
	
	public function getHeaders(includeRules)
		set getHeaders = getMessageHeaders(includeRules, me.MessageDepth, me.CustomOrdering, me.IncludeDetails, me.Fisses)
	end function
	
	private function getTypesHeaders()
		set getTypesHeaders = getMessageTypesHeaders(me.IncludeDetails)
	end function 
	
	Public function getUnifiedMessageTypes()
		set getUnifiedMessageTypes = getMyMessageTypes(true)
	end function
	
	private function getMyMessageTypes(unified)
		dim types
		set types = CreateObject("System.Collections.ArrayList")
		'add base types
		dim baseTypeName
		dim baseTypeElement
		dim elementOrder
		elementOrder = 0
		for each baseTypeName in me.BaseTypes.Keys
			elementOrder = elementOrder + 1
			set baseTypeElement = me.BaseTypes.Item(baseTypeName)
			if not IsObject(baseTypeElement) then
				set baseTypeElement = nothing
			end if
			if not baseTypeElement is nothing then
				if baseTypeElement.Stereotype <> "BDT" then
					'first add the properties for the base type itself
					dim baseTypeProperties
					set baseTypeProperties = getBaseTypeProperties(baseTypeElement)
					if unified then
						'add the messageName
						baseTypeProperties.Insert 0, me.Name
					end if
					'add properties to list
					types.add baseTypeProperties
				end if
			end if
		next
		'add enumerations
		dim enumName
		dim enumElement as EA.Element
		for each enumName in me.Enumerations.Keys
			elementOrder = elementOrder + 1
			set enumElement = me.Enumerations.Item(enumName)
			'add all the literal values
			dim enumLiteral as EA.Attribute
			'if the enum has no values then ALL values are allowed
			'so we look for the enum this enum was based on
			if enumElement.Attributes.Count = 0 then
				'replace enumElement with base enum
				set enumElement = getBaseEnum(enumElement)
			end if
			'loop the enum literals
			for each enumLiteral in enumElement.Attributes
				elementOrder = elementOrder + 1
				dim enumLiteralProperties
				set enumLiteralProperties = getEnumLiteralProperties(enumElement,enumLiteral)
				if unified then
					'add the messageName
					enumLiteralProperties.Insert 0, me.Name
				end if
				types.add enumLiteralProperties
			next
		next
		'return the types
		set getMyMessageTypes = types
	end function
	
	private function getBaseEnum(enumElement)
		'initialize
		set getBaseEnum = nothing
		dim sqlGetBaseEnum
		sqlGetBaseEnum = "select o.Object_ID from t_object o                             " & _
						" inner join t_connector c on c.End_Object_ID = o.Object_ID      " & _
						" 							and c.Connector_Type = 'Abstraction' " & _
						" 							and c.Stereotype = 'trace'           " & _
						" where o.Object_Type = 'Enumeration'                            " & _
						" and c.Start_Object_ID = " & enumElement.ElementID & "          "
		dim baseEnums
		dim baseEnum as EA.Element
		set baseEnums = getElementsFromQuery(sqlGetBaseEnum)
		for each baseEnum in baseEnums
			set getBaseEnum = baseEnum
			exit for 'we only need the first one
		next
	end function
	
	Public function getMessageTypes()
		dim types
		set types = CreateObject("System.Collections.ArrayList")
		'first add the headers
		dim typeHeaders
		set typeHeaders = getTypesHeaders()
		'Session.Output typeHeaders.Count
		types.add typeHeaders
		'get the actual content
		types.AddRange getMyMessageTypes(false)
		'return the types
		set getMessageTypes = types
	end function
	
	Private function getEnumLiteralProperties(enumElement,enumLiteral)
		dim enumLiteralProperties 
		set enumLiteralProperties = CreateObject("System.Collections.ArrayList")
		'first fill the array with empty strings
		fillArrayList enumLiteralProperties, "", 6
		'category
		enumLiteralProperties(0) = "Enumeration"
		'Type
		enumLiteralProperties(1) = enumElement.Name
		'Code
		enumLiteralProperties(2) = "'" & enumLiteral.Name
		'Description
		enumLiteralProperties(3) = enumLiteral.Alias
		'Get the CodeName tagged value if it exists
		dim codeNameTv as EA.AttributeTag
		for each codeNameTv in enumLiteral.TaggedValues
			if lcase(codeNameTv.Name) = "codename" then
				enumLiteralProperties(3) = codeNameTv.Value
				exit for
			end if
		next
		'return the properties
		set getEnumLiteralProperties = enumLiteralProperties
	end function
	
	Private function getBaseTypeProperties(baseType)
		dim baseTypeProperties 
		set baseTypeProperties = CreateObject("System.Collections.ArrayList")
		'first fill the array with empty strings
		fillArrayList baseTypeProperties, "", 6
		'category
		baseTypeProperties(0) = "BaseType"
		'Type
		baseTypeProperties(1) = baseType.Name
		'Code
		baseTypeProperties(2) = "" 'emtpty for the base type
		'Description
		baseTypeProperties(3) = "" 'emtpty for the base type
		'Restriction Base
		dim derivedFrom
		derivedFrom = getDerivedFrom(baseType)
		baseTypeProperties(4) = derivedFrom
		'Facets
		'add properties based on the tagged values
		dim facetSpecification
		facetSpecification = "" 'initial value
		dim tv as EA.TaggedValue
		for each tv in baseType.TaggedValues
			select case tv.Name
				case "fractionDigits", "length", "maxExclusive", "maxInclusive", "maxLength", "minExclusive","minInclusive","minLength",_
				"pattern","totalDigits","whiteSpace", "enumeration"
					facetSpecification = addFacetSpecification(facetSpecification, tv)
			end select
		next
		baseTypeProperties(5) = facetSpecification
		'return the base type properties
		set getBaseTypeProperties = baseTypeProperties
	end function
	
	private function addFacetSpecification(facetSpecification, facetTV)
		addFacetSpecification = facetSpecification 'initial value
		if len(facetTV.Value) > 0 then
			if len(facetSpecification) > 0  then
				addFacetSpecification = addFacetSpecification & vbNewLine
			end if
			addFacetSpecification = addFacetSpecification & facetTV.Name & ": " & facetTV.Value
		end if
	end function
	
	private function fillArrayList(listToFill, fillValue, count)
		dim i
		for i = 0 to count -1 step +1
			listToFill.Add fillValue
		next
	end function
	
	private function getDerivedFrom(baseType)
		'the base type either inherits from a standard XSD type, or has it stored separately (gentype?)
		dim baseTypeBaseTypes 
		set baseTypeBaseTypes = baseType.BaseClasses
		dim derivedFrom as EA.Element
		set derivedFrom = nothing
		'get the first base class
		for each derivedFrom in baseType.BaseClasses
			exit for
		next
		if not IsObject(derivedFrom) then
			set derivedFrom = nothing
		end if
		'Check for CON attribute of BDT element
		if derivedFrom is nothing _
		and baseType.Stereotype = "BDT" then
			'get CON attributre
			dim attribute as EA.Attribute
			for each attribute in baseType.Attributes
				if attribute.Stereotype = "CON" _
				and attribute.ClassifierID > 0 then
					set derivedFrom = Repository.GetElementByID(attribute.ClassifierID)
				end if
			next
		end if
		'set name if derivedFrom element is found
		if not derivedFrom is nothing then
			getDerivedFrom = derivedFrom.Name
		else
			'if there is no real inheritance link then the link is stored in the genLinks property as parent=<name>;
			getDerivedFrom = getValueForkey(baseType.Genlinks, "parent")
		end if
		if getDerivedFrom = "anySimpleType" then
			'we are not interested in "anySimpletype"
			getDerivedFrom = ""
		end if
	end function
	
end Class

'"Static" functions

public function getMessageHeaders(includeRules, depth, customOrdering, technical, fisses)
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	'first order
	headers.add("Order")
	'then Message
	headers.Add("Message")
	'add the levels
	dim i
	for i = 1 to depth -1 step +1
		headers.add("L" & i)
	next
	'Cardinality
	headers.Add("Cardinality")
	'Type
	headers.Add("Type")
	'base type
	headers.Add("Base Type")
	'Constraints (facets)
	headers.Add("Constraints")
	if customOrdering then
		if not technical then
			'LDM mapping
			'LDM Class
			headers.Add("LDM Class")
			'LDM Attribute
			headers.Add("LDM Attribute")
			'BusinessEntity
			headers.Add("Business Entity")
			'BusinessAttribute
			headers.Add("Business Attribute")
		end if
	elseif includeRules then
		'with our without test rules
		'Test Rule ID
		headers.Add("Test Rule ID")
		'Test Rule
		headers.Add("Test Rule")
		'Error Reason
		headers.Add("Error Reason")
	end if
	'add business usage
	if customOrdering _
	and not technical _
	and not fisses is nothing then
		dim fis as EA.Element
		for each fis in fisses
			headers.Add fis.Name
		next
	end if
	'return the headers
	set getMessageHeaders = headers
end function

private function getMessageTypesHeaders(unified)
		dim headers
		set headers = CreateObject("System.Collections.ArrayList")
		'Message
		if unified then
			headers.Add("Message")
		end if
		'Category
		headers.Add("Category") 'Enumeration or BaseType '0
		'Type
		headers.Add("Type") '1
		'Code
		headers.Add("Code") '2
		'Description
		headers.Add("Description") '3
		'Restriction Base
		headers.Add("Restriction Base") '4
		'Facets
		headers.Add("Facets") '5
		'return the headers
		set getMessageTypesHeaders = headers
end function