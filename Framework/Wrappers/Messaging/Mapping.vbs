'[path=\Framework\Wrappers\Messaging]
'[group=Messaging]

!INC Utils.Include

' Author: Geert Bellekens
' Purpose: A wrapper class for a mapping
' Date: 2019-03-08


'constants
const linkedAttributeTag = "linkedAttribute"
const linkedAssociatonTag = "linkedAssociation"
const linkedElementTag = "sourceElement"

Class Mapping

	
	'private variables
	Private m_Target
	private m_TaggedValue
	private m_IsEmpty
	private m_MappingLogics
	private m_MappingPathString
	private m_TargetParent
	
	'constructor
	Private Sub Class_Initialize
		set m_Target = nothing
		set m_TaggedValue = nothing
		m_IsEmpty = false
		set m_MappingLogics = CreateObject("System.Collections.ArrayList")
		m_MappingPathString = ""
		set m_TargetParent = nothing
	End Sub
	
	'public properties
	
	' Target property.
	Public Property Get Target
	  set Target = m_Target
	End Property
	
	' TaggedValue property.
	Public Property Get TaggedValue
		set TaggedValue = m_TaggedValue
	End Property
	Public Property Let TaggedValue(value)
		set m_TaggedValue = value
		'get the details from the tagged value
		setMappingDetails
	End Property
	
	'MappingPathString property.
	Public Property Get MappingPathString
		MappingPathString = m_MappingPathString
	End Property
	
	'IsEmpty property.
	Public Property Get IsEmpty
		IsEmpty = m_IsEmpty
	End Property
	
	' MappingLogics property.
	Public Property Get MappingLogics
		set MappingLogics = m_MappingLogics
	End Property
	Public Property Let MappingLogics(value)
		set m_MappingLogics = value
	End Property
		
	' TargetParent property.
	Public Property Get TargetParent
	  set TargetParent = m_TargetParent
	End Property
	
	private function setMappingDetails()
		if m_TaggedValue is nothing then
			exit function
		end if
		'get target
		if lcase(m_TaggedValue.Name) = lcase(linkedAttributeTag) then
			set m_Target = Repository.GetAttributeByGuid(m_TaggedValue.Value)
		elseif lcase(m_TaggedValue.Name) = lcase(linkedAssociatonTag) then
			set m_Target = Repository.GetConnectorByGuid(m_TaggedValue.Value)
		elseif lcase(m_TaggedValue.Name) = lcase(linkedElementTag) then
			set m_Target = Repository.GetElementByGuid(m_TaggedValue.Value)
		end if
		'get the details from the notes
		if len(m_TaggedValue.Notes) = 0 then
			exit function
		end if
		Dim xDoc 
		Set xDoc = CreateObject("Microsoft.XMLDOM")
		'Set xDoc = CreateObject("Msxml2.DOMDocument")
		'get the mappingPath
		xDoc.LoadXML m_TaggedValue.Notes
		dim sourcePathNode 
		set sourcePathNode = xDoc.SelectSingleNode("//mappingSourcePath")
		if not sourcePathNode is nothing then
			m_MappingPathString = sourcePathNode.Text
		end if
		'get the target mapping path
		dim targetPathNode 
		set targetPathNode = xDoc.SelectSingleNode("//mappingTargetPath")
		if not targetPathNode is nothing then
			dim targetNodes
			targetNodes = Split(targetPathNode.Text, ".")
			if Ubound(targetNodes) >= 1 then
				'get the second last guid
				dim parentGUID 
				parentGUID = targetNodes(Ubound(targetNodes) -1)
				set m_TargetParent = Repository.GetElementByGuid(parentGUID)
			end if
		end if
		'get the IsEmpty Property
		dim isEmptyNode
		set isEmptyNode = xDoc.SelectSingleNode("//isEmptyMapping")
		if not isEmptyNode is nothing then
			if lcase(isEmptyNode.Text) = "true" then
				m_IsEmpty = true
			end if
		end if
		'get the mapping logics
		dim mappingLogicNodes
		set mappingLogicNodes = xDoc.SelectNodes("//mappingLogic")
		dim mappingLogicNode
		for each mappingLogicNode in mappingLogicNodes
			'create mapping logic
			dim mappingLogic
			set mappingLogic = new MappingLogic
			dim contextNode
			set contextNode = mappingLogicNode.SelectSingleNode("./context")
			if not contextNode is nothing then
				mappingLogic.Context = Repository.GetElementByGuid(contextNode.Text)
			end if
			dim descriptionNode
			set descriptionNode = mappingLogicNode.SelectSingleNode("./description")
			if not descriptionNode is nothing then
				mappingLogic.Description = descriptionNode.Text
			end if
			if not mappingLogic.Context is nothing or len(mappingLogic.Description) > 0 then
				m_MappingLogics.Add mappingLogic
			end if
		next
	end function
end Class