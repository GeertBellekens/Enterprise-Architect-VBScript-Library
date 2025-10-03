'[path=\Framework\YAML]
'[group=YAML]

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: YamlScalarNode
' Author: Geert Bellekens
' Purpose: Class for Yaml Scalar Nodes (having a primitive value)
' Date: 2025-03-18
'

'scalarNode has a "string" value
Class YamlScalarNode
	private  m_scalarValue
	private m_parentNode
	private m_SubNodes
	
	Private Sub Class_Initialize
		m_scalarValue = ""
		set m_parentNode = nothing
		set m_SubNodes = nothing
	end sub
	
	public Property Get Indentation
		dim myIndent
		if me.Level < 2 then
			myIndent = ""
		else
			myIndent = me.ParentNode.Indentation & space(2)
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
		
	Public Property Get ParentNode
		set ParentNode = m_ParentNode
	End Property
	Public Property Let ParentNode(value)
		set m_ParentNode = value
	End Property
	
	Public Property Get ScalarValue
		ScalarValue = m_scalarValue
	End Property
	Public Property Let ScalarValue(value)
		m_scalarValue = value
		'strip quotes
	End Property
	Public function ToString(noIndent)
		ToString = replace(me.ScalarValue, vbNewLine,  vbNewLine & me.Indentation )
	end function
	public function find(keyName, recurse)
		'scalar nodes can never be found as they don't have a keyname
		set find = nothing
	end function
	Public Property Get SubNodes
		if m_SubNodes is nothing then
			set m_SubNodes = CreateObject("System.Collections.ArrayList")
		end if
		set SubNodes = m_SubNodes
	end property
	
	public function getNodeByName(keyToFind)
		dim node
		set node = nothing
		set getNodeByName = node
	end function
end class