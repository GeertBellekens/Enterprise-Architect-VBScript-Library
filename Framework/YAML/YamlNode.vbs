'[path=\Framework\YAML]
'[group=YAML]
'option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: YAML
' Author: Geert Bellekens
' Purpose: Parse YAML file in a tree structure
' Date: 2024-12-24
'
class YamlNode
	private m_key
	private m_value
	private m_parentNode
	private m_IsSequence
	private m_SubNodes
	
	Private Sub Class_Initialize
		set m_key = nothing
		set m_value = nothing
		set m_parentNode = nothing
		m_IsSequence = false
		set m_SubNodes = nothing
	end sub
	
	public Property Get Indentation
		dim myIndent
		if me.Level < 2 then
			myIndent = ""
		else
			myIndent = me.ParentNode.Indentation
			if not me.IsSequence then
				myIndent = myIndent & space(2)
			end if
		end if
		'return
		Indentation = myIndent
	end Property
	
	public Property Get Level
		if not me.ParentNode is nothing then
			Level = me.ParentNode.Level + 1
		else
			Level = 0
		end if
	end Property
	
	Public Property Get Key
		set Key = m_Key
	End Property
	Public Property Let Key(value)
		set m_Key = value
		m_Key.ParentNode = me
	End Property
	public Property Get KeyName
		KeyName = ""
		if not me.Key is nothing then
			KeyName = me.Key.ScalarValue
		end if
	End Property
	Public Property Get SubNodes
		if m_SubNodes is nothing then
			if not me.Value is nothing then
				set m_SubNodes = me.Value.SubNodes
			else
				set m_SubNodes = CreateObject("System.Collections.ArrayList")
			end if
		end if
		set SubNodes = m_SubNodes
	end Property
	
	Public Property Get Value
		set Value = m_value
	End Property
	Public Property Let Value(in_value)
		set m_value = in_value
		m_value.ParentNode = me
	End Property
	Public Property Get ValueName
		ValueName = ""
		if not me.Value is nothing then
			if TypeName(me.Value) = "YamlScalarNode" then
				ValueName = me.Value.ScalarValue
			end if
		end if
	End Property
	Public Property Get IsSequence
		IsSequence = m_IsSequence
	End Property
	Public Property Let IsSequence(value)
		m_IsSequence = value
	End Property
	
	Public Property Get ParentNode
		set ParentNode = m_ParentNode
	End Property
	Public Property Let ParentNode(value)
		set m_ParentNode = value
	End Property
	
	Public function getParent(requestedIndent)
		if len(requestedIndent) <= len(me.Indentation) _
		  and not me.ParentNode is nothing then
			set getParent = me.ParentNode.getParent(requestedIndent)
		else
			set getParent = me.Value
		end if
	end function
	
	Public function ToString(noIndent)
		dim indent
		indent = ""
		'we need to add a space if the value ToString() doesn't start with a newline
		dim spacer
		spacer = ""
		dim valueString
		valueString = ""
		if not me.Value is nothing then
			valueString = me.Value.ToString(false)
		end if
		if not (left(valueString, len(vbNewLine))) = vbNewLine then
			spacer = space(1)
		end if
		if me.IsSequence then
			ToString = me.Indentation & "-" & spacer & valueString
		else
			if not noIndent then
				indent = me.Indentation
			end if
			ToString = indent & me.Key.ToString(false) & ":" & spacer & valueString
		end if
	end function
	
	public function setScalarKey(keyValue)
		dim keyNode
		set keyNode = new YamlScalarNode
		keyNode.ScalarValue = keyValue
		me.Key = keyNode
	end function
	
	public function setScalarValue(valueValue)
		dim valueNode
		set valueNode = new YamlScalarNode
		valueNode.ScalarValue = valueValue
		me.Value = valueNode
	end function
	
	public function appendScalarValue(valueValue)
		if me.Value is nothing then
			me.setScalarValue("")
		end if
		me.Value.ScalarValue = me.Value.ScalarValue & valueValue
	end function
	
		
	public function find(keyToFind, recurse)
		'return this node if the keyname corresponds to the name of the key node
		set find = nothing
		if lcase(me.KeyName) = lcase(keyToFind) then
			set find = me
			exit function
		end if
		if not me.Value is nothing _
		  and recurse then
			set find = me.Value.find(keyToFind, recurse)
		end if
	end function
	
	public function getValueName(keyToFind)
		dim valueName
		valueName = "" 'default
		dim node
		'set node = me.find(keyToFind)
		if not me.Value is nothing then
			set node = me.Value.find(keyToFind, false)
		end if
		if not node is nothing then
			valueName = node.ValueName
		end if
		'return
		getValueName = valueName
	end function
	
	public function getNodeByName(keyToFind)
		dim node
		set node = nothing
		'create parentNode if needed
		if me.Value is nothing then
			me.Value = new YamlParentNode
		end if
		'get the node
		set node = me.Value.getNodeByName(keyToFind)
		'return
		set getNodeByName = node
	end function
	
	public function addKeyValueNode(key, value)
		if len(value) = 0 then
			exit function
			'TODO check if node exists, and then remove it
		end if
		dim node
		set node = me.GetNodeByName(key)
		node.setScalarValue(value)
		'return
		set addKeyValueNode = node
	end function
	
end class





