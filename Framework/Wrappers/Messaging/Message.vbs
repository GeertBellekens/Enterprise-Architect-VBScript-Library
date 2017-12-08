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

	'constructor
	Private Sub Class_Initialize
		m_Name = ""
		set m_RootNode = nothing
		m_MessageDepth = 0
		set m_BaseTypes = CreateObject("Scripting.Dictionary")
		set m_Enumerations = CreateObject("Scripting.Dictionary")
		m_Prefix = ""
		set m_ValidationRules = CreateObject("System.Collections.ArrayList")
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
		
		'with our without test rules
		if includeRules then
			'Test Rule ID
			headers.Add("Test Rule ID")
			'Test Rule
			headers.Add("Test Rule")
			'Error Reason
			headers.Add("Error Reason")
		end if
		
		'return the headers
		set getHeaders = headers
	end function
	
	private function getTypesHeaders()
		dim headers
		set headers = CreateObject("System.Collections.ArrayList")
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
		'fractionDigits
		headers.Add("fractionDigits") '5
		'length
		headers.Add("length") '6
		'maxExclusive
		headers.Add("maxExclusive") '7
		'maxInclusive
		headers.Add("maxInclusive") '8
		'maxLength
		headers.Add("maxLength")'9
		'minExclusive
		headers.Add("minExclusive") '10
		'minInclusive
		headers.Add("minInclusive") '11
		'minLength
		headers.Add("minLength") '12
		'pattern
		headers.Add("pattern") '13
		'totalDigits
		headers.Add("totalDigits") '14
		'whiteSpace
		headers.Add("whiteSpace") '15
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
			if not IsObject(baseTypeElement) then
				set baseTypeElement = nothing
			end if
			if not baseTypeElement is nothing then
				'first add the properties for the base type itself
				dim baseTypeProperties
				set baseTypeProperties = getBaseTypeProperties(baseTypeElement)
				types.add baseTypeProperties
			end if
		next
		'add enumerations
		dim enumName
		dim enumElement
		for each enumName in me.Enumerations.Keys
			elementOrder = elementOrder + 1
			set enumElement = me.Enumerations.Item(enumName)
			'add all the literal values
			dim enumLiteral as EA.Attribute
			for each enumLiteral in enumElement.Attributes
				elementOrder = elementOrder + 1
				dim enumLiteralProperties
				set enumLiteralProperties = getEnumLiteralProperties(enumElement,enumLiteral)
				types.add enumLiteralProperties
			next
		next
		'return the types
		set getMessageTypes = types
	end function
	
	Private function getEnumLiteralProperties(enumElement,enumLiteral)
		dim enumLiteralProperties 
		set enumLiteralProperties = CreateObject("System.Collections.ArrayList")
		'first fill the array with empty strings
		fillArrayList enumLiteralProperties, "", 16
		'category
		enumLiteralProperties(0) = "Enumeration"
		'Type
		enumLiteralProperties(1) = enumElement.Name
		'Code
		enumLiteralProperties(2) = enumLiteral.Name
		'Description
		enumLiteralProperties(3) = enumLiteral.Alias
		'return the properties
		set getEnumLiteralProperties = enumLiteralProperties
	end function
	
	Private function getBaseTypeProperties(baseType)
		dim baseTypeProperties 
		set baseTypeProperties = CreateObject("System.Collections.ArrayList")
		'first fill the array with empty strings
		fillArrayList baseTypeProperties, "", 16
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
		'add properties based on the tagged values
		dim tv as EA.TaggedValue
		for each tv in baseType.TaggedValues
			select case tv.Name
				case "fractionDigits"
					baseTypeProperties(5) = tv.Value'5
				case "length"
					baseTypeProperties(6) = tv.Value '6
				case "maxExclusive"
					baseTypeProperties(7) = tv.Value '7
				case "maxInclusive"
					baseTypeProperties(8) = tv.Value '8
				case "maxLength"
					baseTypeProperties(9) = tv.Value '8
				case "minExclusive"
					baseTypeProperties(10) = tv.Value '10
				case "minInclusive"
					baseTypeProperties(11) = tv.Value '11
				case "minLength"
					baseTypeProperties(12) = tv.Value '12
				case "pattern"
					baseTypeProperties(13) = tv.Value '13
				case "totalDigits"
					baseTypeProperties(14) = tv.Value '14
				case "whiteSpace"
					baseTypeProperties(15) = tv.Value'15
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
		if not IsObject(derivedFrom) then
			set derivedFrom = nothing
		end if
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