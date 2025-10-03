'[path=\Framework\YAML]
'[group=YAML]

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: YamlScalarNode
' Author: Geert Bellekens
' Purpose: Class for Yaml Scalar Nodes (having a primitive value)
' Date: 2025-03-18
'


Class YamlParentNode
	private m_mappingNodes 'MappingNode has an unordered list of nodes with unique keys
	private m_sequenceNodes 'Sequence nodes is an ordered list of nodes
	private m_parentNode
	private m_noIndent
	private m_SubNodes

	
	Private Sub Class_Initialize
		set m_mappingNodes = CreateObject("Scripting.Dictionary")
		set m_sequenceNodes = CreateObject("System.Collections.ArrayList")
		set m_parentNode = nothing
		set m_SubNodes = nothing
		m_noIndent = false
	end sub
	
	public Property Get Indentation
		dim myIndent
		if me.Level < 2 then
			myIndent = ""
		else
			myIndent = me.ParentNode.Indentation
		end if
		'return
		Indentation = myIndent
	end Property
	
	public Property Get Level
		if not me.ParentNode is nothing then
			Level = me.ParentNode.Level
		else
			Level = 0
		end if
	end Property
	
	Public Property Get ParentNode
		set ParentNode = m_ParentNode
	End Property
	Public Property Let ParentNode(value)
		set m_ParentNode = value
	End Property
	
	Public Property Get NoIndent
		NoIndent = m_noIndent
	End Property
	Public Property Let NoIndent(value)
		m_noIndent = value
	End Property
	
	public function addNode(yamlNode)
		if yamlNode.IsSequence then
			m_sequenceNodes.Add yamlNode
		else
			set m_mappingNodes(yamlNode.Key) = yamlNode
		end if
		yamlNode.ParentNode = me
	end function
	

	
	Public function ToString(noIndent)
		dim returnString
		returnString = ""
		dim node
		for each node in m_mappingNodes.items
			'the first item of the list should not have an indent if "NoIndent" is on
			if me.NoIndent and len(returnString) = 0 then
				'add content without indent
				returnString = returnString & node.ToString(true)
			else
				'add newline if needed (not needed for the first top level node)
				if not (len(returnString) = 0 and me.level = 0) then
					returnString = returnString & vbNewLine
				end if
				'add content
				returnString = returnString & node.ToString(false)
			end if
		next
		for each node in m_sequenceNodes
			'add newline if needed (not needed for the first top level node)
			if not (len(returnString) = 0 and me.level = 0) then
				returnString = returnString & vbNewLine
			end if
			'add content
			returnString = returnString & node.ToString(false)
		next
		'return
		ToString = returnString
	end function
	Public function getParent(requestedIndent)
		if len(requestedIndent) <= len(me.Indentation) _
		  and not me.ParentNode is nothing then
			set getParent = me.ParentNode.getParent(requestedIndent)
		else
			set getParent = me
		end if
	end function
	public function find(keyName, recurse)
		set find = nothing
		dim node
		'first look in all subnodes
		for each node in me.SubNodes
			set find = node.find(keyName, recurse)
			if not find is nothing then
				exit function
			end if
		next
	end function
	Public Property Get SubNodes
		if m_SubNodes is nothing then
			set m_SubNodes = CreateObject("System.Collections.ArrayList")
			dim node
			for each node in m_mappingNodes.items
				m_SubNodes.Add node
			next
			for each node in m_sequenceNodes
				m_SubNodes.Add node
			next
		end if
		set SubNodes = m_SubNodes
	end property
	
	public function getNodeByName(keyToFind)
		dim node
		set node = nothing
		dim currentNode
		'find existing one
		for each currentNode in me.SubNodes
			if currentNode.KeyName = keyToFind then
				set node = currentNode
			end if
		next
		'create new if needed
		if node is nothing then
			set node = new YamlNode
			node.setScalarKey(keyToFind)
			me.AddNode(node)
		end if
		set getNodeByName = node
	end function
end class