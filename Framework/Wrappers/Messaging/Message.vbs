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
	
	public function loadMessage(eaRootNodeElement)
		'set the name of the message
		
		'the name of the message is equal to the name of the owning package, unless the rootnodeElement is a package
		dim ownerPackage as EA.Package
		if eaRootNodeElement.ObjectType = otPackage then
			set ownerPackage = eaRootNodeElement
		else
			set ownerPackage = Repository.GetPackageByID(eaRootNodeElement.PackageID)
		end if
		me.Name = eaRootNodeElement.Name
		'set the prefix
		m_Prefix = getPrefix(ownerPackage)
		'set the customOrdering property (check if ï¿½MAï¿½ is one of the stereotypes
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
		me.RootNode.intitializeWithSource eaRootNodeElement, nothing, "1..1", nothing, nothing
		'if the rootnode has no subnodes and it has a nested complex type then we initialize it with that complex type
		if me.RootNode.ChildNodes.Count = 0 then
			dim nestedClass as EA.Element
			for each nestedClass in eaRootNodeElement.Elements
				if nestedClass.Stereotype = "XSDcomplexType" then
					me.RootNode.intitializeWithSource nestedClass, nothing, "1..1", nothing, nothing
					me.RootNode.Name = eaRootNodeElement.Name
				end if
			next
		end if
		'set base types and enumerations
		setBaseTypesAndEnumerations(me.RootNode)
		'sort enumerations
		sortEnumerations
		'link the message validation rules
		getMessageValidationRules()
	end function
	
	private function sortEnumerations
		dim sortedKeys 
		set sortedKeys = CreateObject("System.Collections.ArrayList")
		dim key
		for each key in me.Enumerations.Keys
			sortedKeys.Add key
		next
		'sort the keys alphabetically
		sortedKeys.Sort
		'create new dictionary in the correct order
		dim newEnumDic
		set newEnumDic = CreateObject("Scripting.Dictionary")
		for each key in sortedKeys
			newEnumDic.Add key, m_Enumerations(key)
		next
		'set the new dictionary
		set m_Enumerations = newEnumDic
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
	public function createOutput(includeRules)
		dim outputList
		'create empty list for current path
		dim currentPath
		set currentPath = CreateObject("System.Collections.ArrayList")
		'start with the rootnode
		set outputList = me.RootNode.getOutput(1,currentPath,me.MessageDepth, includeRules)
		'return outputlist
		set createOutput = outputList
	end function
	
	public function createStructureOutput(exportType)
		dim outputList
		'create empty list for current path
		dim currentPath
		set currentPath = CreateObject("System.Collections.ArrayList")
		'start with the rootnode
		set outputList = me.RootNode.getStructureOutput(1,currentPath,me.MessageDepth, exportType)
		'return outputlist
		set createStructureOutput = outputList
	end function
	
	
	'create an arraylist of arraylists with the details of this message
	public function createUnifiedOutput(includeRules, depth)
		dim outputList
		'create empty list for current path
		dim currentPath
		set currentPath = CreateObject("System.Collections.ArrayList")
		'start with the rootnode
		set outputList = me.RootNode.getOutput(1,currentPath,depth, includeRules)
		'return outputlist
		set createUnifiedOutput = outputList
	end function
	
	'create an arraylist of arraylists with the details of this message including he headers
	public function createFullOutput(includeRules)
		dim fullOutput
		dim headers
		set fullOutput = me.createOutput(includeRules)
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
		set getHeaders = getMessageHeaders(includeRules, me.MessageDepth, me.CustomOrdering, me.IncludeDetails)
	end function
		
	private function getMyMessageTypes()
		dim types
		set types = CreateObject("System.Collections.ArrayList")
		'add base types
		dim baseTypeName
		dim baseTypeElement
		'add enumerations
		dim enumName
		dim enumElement as EA.Element
		'add the enumeration
		for each enumName in me.Enumerations.Keys
			set enumElement = me.Enumerations.Item(enumName)
			'add the enumeration itself
			types.add getEnumProperties(enumElement)
			'add all the literal values
			dim enumLiteral as EA.Attribute
			'loop the enum literals
			for each enumLiteral in enumElement.Attributes
				dim enumLiteralProperties
				set enumLiteralProperties = getEnumLiteralProperties(enumElement,enumLiteral)
				types.add enumLiteralProperties
			next
			'add an empty row
			dim emptyRow 
			set emptyRow = CreateObject("System.Collections.ArrayList")
			'first fill the array with empty strings
			fillArrayList emptyRow, "", 4
			types.add emptyRow
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
		'get the actual content
		types.AddRange getMyMessageTypes()
		'return the types
		set getMessageTypes = types
	end function
	
	Private function getEnumProperties(enumElement)
		dim enumProperties 
		set enumProperties = CreateObject("System.Collections.ArrayList")
		'first fill the array with empty strings
		fillArrayList enumProperties, "", 4
		'Type
		enumProperties(0) = enumElement.Name
		'Description
		enumProperties(2) = enumElement.Alias & " " & Repository.GetFormatFromField("TXT", enumElement.Notes)
		'Extra functional information
		dim tv as EA.TaggedValue
		for each tv in enumElement.TaggedValues
			if tv.Name = "remark API-Portal" then
				enumProperties(3) = tv.Value
				exit for
			end if
		next
		'return the properties
		set getEnumProperties = enumProperties
	end function
	
	Private function getEnumLiteralProperties(enumElement,enumLiteral)
		dim enumLiteralProperties 
		set enumLiteralProperties = CreateObject("System.Collections.ArrayList")
		'first fill the array with empty strings
		fillArrayList enumLiteralProperties, "", 4
		'Code
		enumLiteralProperties(1) = "'" & enumLiteral.Name
		'Description
		enumLiteralProperties(2) = enumLiteral.Alias & " " & Repository.GetFormatFromField("TXT", enumLiteral.Notes)
		'Extra functional information
		dim tv as EA.AttributeTag
		for each tv in enumLiteral.TaggedValues
			if tv.Name = "remark API-Portal" then
				enumLiteralProperties(3) = tv.Value
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

public function getMessageHeaders(includeRules, depth, customOrdering, technical)
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	'first order
	headers.add("Order")
	'level
	headers.add("Level")
	'then Message
	headers.Add("Message")
	'Element Name
	headers.Add("Name")
	'Description
	headers.Add("Description")
	'Cardinality
	headers.Add("Cardinality")
	'Type
	headers.Add("Type")
	'base type
	headers.Add("Base Type")
	'Facets
	headers.Add("Facets")
	
	if includeRules then
		'with our without test rules
		'Test Rule ID
		headers.Add("Test Rule ID")
		'Test Rule
		headers.Add("Test Rule")
		'Error Reason
		headers.Add("Error Reason")
	end if
	'return the headers
	set getMessageHeaders = headers
end function

