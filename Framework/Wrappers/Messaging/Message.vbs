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

	'constructor
	Private Sub Class_Initialize
		m_Name = ""
		set m_RootNode = nothing
		m_MessageDepth = 0
		set m_BaseTypes = CreateObject("Scripting.Dictionary")
		set m_Enumerations = CreateObject("Scripting.Dictionary")
		m_Prefix = ""
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
	
	public function loadMessage(eaRootNodeElement)
		'set the name of the message
		'the name of the message is equal to the name of the owning package
		dim ownerPackage as EA.Package
		set ownerPackage = Repository.GetPackageByID(eaRootNodeElement.PackageID)
		me.Name = ownerPackage.Name
		'set the prefix
		m_Prefix = getPrefix(ownerPackage)
		'create the root node
		me.RootNode = new MessageNode
		me.RootNode.intitializeWithSource eaRootNodeElement, nothing, "1..1", nothing, nothing
		setBaseTypesAndEnumerations(me.RootNode)
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
			if not foundEnumeration _
			AND not me.BaseTypes.Exists(messageNode.TypeName) then
				'add to BaseTypes list
				me.BaseTypes.Add messageNode.TypeName, messageNode.TypeElement
			end if
		else
			'not a leafnode, check the childnodes
			dim childNode
			for each childNode in messageNode.ChildNodes
				setBaseTypesAndEnumerations childNode 
			next
		end if
	end function
	
	'create an arraylist of arraylists with the details of this message
	public function createOuput()
		dim outputList
		'create empty list for current path
		dim currentPath
		set currentPath = CreateObject("System.Collections.ArrayList")
		'start with the rootnode
		set outputList = me.RootNode.getOuput(1,currentPath,me.MessageDepth)
		'return outputlist
		set createOuput = outputList
	end function
	
	
	'gets the maximum depth of this message
	private function getMessageDepth()
		dim message_depth
		message_depth = 0
		message_depth = me.RootNode.getDepth(message_depth)
		getMessageDepth = message_depth
	end function
	
	public function getHeaders()
		dim headers
		set headers = CreateObject("System.Collections.ArrayList")
		'first order
		headers.add("Order")
		'then Message
		headers.Add("Message")
		'add the levels
		dim i
		for i = 1 to me.MessageDepth -1 step +1
			headers.add("L" & i)
		next
		'Cardinality
		headers.Add("Cardinality")
		'Type
		headers.Add("Type")
		'Test Rule
		headers.Add("Test Rule")
		'Test Rule ID
		headers.Add("Test Rule ID")
		'Error Reason
		headers.Add("Error Reason")
		'return the headers
		set getHeaders = headers
	end function
	
	private function getTypesHeaders()
		dim headers
		set headers = CreateObject("System.Collections.ArrayList")
		'first order
		headers.add("Order") '0
		'Category
		headers.Add("Category") 'Enumeration or BaseType '1
		'Type
		headers.Add("Type") '2
		'Code
		headers.Add("Code") '3
		'Description
		headers.Add("Description") '4
		'Derivation
		headers.Add("Derivation") 'Restriction or List '5
		'DerivedFrom
		headers.Add("Derived From") '6
		'fractionDigits
		headers.Add("FractionDigits") '7
		'length
		headers.Add("Length") '8
		'maxExclusive
		headers.Add("MaxExclusive") '9
		'maxInclusive
		headers.Add("MaxInclusive") '10
		'maxLength
		headers.Add("MaxLength")'11
		'minExclusive
		headers.Add("MinExclusive") '12
		'minInclusive
		headers.Add("MinInclusive") '13
		'minLength
		headers.Add("MinLength") '14
		'pattern
		headers.Add("Pattern") '15
		'totalDigits
		headers.Add("TotalDigits") '16
		'whiteSpace
		headers.Add("WhiteSpace") '17
		'return the headers
		set getTypesHeaders = headers
	end function
	
	Public function getMessageTypes()
		dim types
		set types = CreateObject("System.Collections.ArrayList")
		'first add the headers
		dim typeHeaders
		set typeHeaders = getTypesHeaders()
		types.add typeHeaders
		'add base types
		dim baseTypeName
		dim baseTypeElement
		dim elementOrder
		elementOrder = 0
		for each baseTypeName in me.BaseTypes.Keys
			elementOrder = elementOrder + 1
			set baseTypeElement = me.BaseTypes.Item(baseTypeName)
			'first add the properties for the base type itself
			dim baseTypeProperties
			set baseTypeProperties = getBaseTypeProperties(baseTypeElement,elementOrder)
			types.add baseTypeProperties
		next
		'add enumerations
		dim enumName
		dim enumElement
		for each enumName in me.Enumerations.Keys
			elementOrder = elementOrder + 1
			set enumElement = me.Enumerations.Item(enumName)
			'first add the properties for the enum itself
			dim enumProperties
			set enumProperties = getEnumProperties(enumElement,elementOrder)
			types.add enumProperties
			'then add all the literal values
			dim enumLiteral as EA.Attribute
			for each enumLiteral in enumElement.Attributes
				elementOrder = elementOrder + 1
				dim enumLiteralProperties
				set enumLiteralProperties = getEnumLiteralProperties(enumElement,enumLiteral,elementOrder)
				types.add enumLiteralProperties
			next
		next
		'return the types
		set getMessageTypes = types
	end function
	
	Private function getEnumLiteralProperties(enumElement,enumLiteral,elementOrder)
		dim enumLiteralProperties 
		set enumLiteralProperties = CreateObject("System.Collections.ArrayList")
		'first fill the array with empty strings
		fillArrayList enumLiteralProperties, "", 18
		'order
		enumLiteralProperties(0) = elementOrder
		'category
		enumLiteralProperties(1) = "Enumeration"
		'Type
		enumLiteralProperties(2) = enumElement.Name
		'Code
		enumLiteralProperties(3) = enumLiteral.Name
		'Description
		enumLiteralProperties(4) = enumLiteral.Alias
		'return the properties
		set getEnumLiteralProperties = enumLiteralProperties
	end function
	
	Private function getEnumProperties(enumElement,elementOrder)
		dim enumProperties 
		set enumProperties = CreateObject("System.Collections.ArrayList")
		'first fill the array with empty strings
		fillArrayList enumProperties, "", 18
		'order
		enumProperties(0) = elementOrder
		'category
		enumProperties(1) = "Enumeration"
		'Type
		enumProperties(2) = enumElement.Name
		'Code
		enumProperties(3) = "" 'emtpty for the enum itself
		'Description
		enumProperties(4) = "" 'emtpty for the enum itself
		'return the properties
		set getEnumProperties = enumProperties
	end function
	
	Private function getBaseTypeProperties(baseType,elementOrder)
		dim baseTypeProperties 
		set baseTypeProperties = CreateObject("System.Collections.ArrayList")
		'first fill the array with empty strings
		fillArrayList baseTypeProperties, "", 18
		'order
		baseTypeProperties(0) = elementOrder
		'category
		baseTypeProperties(1) = "BaseType"
		'Type
		baseTypeProperties(2) = baseType.Name
		'Code
		baseTypeProperties(3) = "" 'emtpty for the base type
		'Description
		baseTypeProperties(4) = "" 'emtpty for the base type
		'derived from
		dim derivedFrom
		derivedFrom = getDerivedFrom(baseType)
		baseTypeProperties(6) = derivedFrom
		'add properties based on the tagged values
		dim tv as EA.TaggedValue
		for each tv in baseType.TaggedValues
			select case tv.Name
				case "derivation"
					baseTypeProperties(5) = tv.Value '5
				case "fractionDigits"
					baseTypeProperties(7) = tv.Value'7
				case "length"
					baseTypeProperties(8) = tv.Value '8
				case "maxExclusive"
					baseTypeProperties(9) = tv.Value '9
				case "maxInclusive"
					baseTypeProperties(10) = tv.Value '10
				case "maxLength"
					baseTypeProperties(11) = tv.Value '11
				case "minExclusive"
					baseTypeProperties(12) = tv.Value '12
				case "minInclusive"
					baseTypeProperties(13) = tv.Value '13
				case "minLength"
					baseTypeProperties(15) = tv.Value '15
				case "pattern"
					baseTypeProperties(15) = tv.Value '15
				case "totalDigits"
					baseTypeProperties(17) = tv.Value '17
				case "whiteSpace"
					baseTypeProperties(18) = tv.Value'18
			end select
		next
		'return the base type properties
		set getBaseTypeProperties = baseTypeProperties
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