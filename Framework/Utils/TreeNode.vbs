'[path=\Framework\Utils]
'[group=Utils]

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: TreeNode
' Author: Geert Bellekens
' Purpose: A class that can be used to represent hiearchical tree structures in EA
' Date: 2025-04-22
'
dim archimateCompositionParentQuery 'not a const because of the & vbNewLine stuff
archimateCompositionParentQuery = "select po.Object_ID from t_object o                                            " & vbNewLine & _
										" inner join t_connector c on c.End_Object_ID = o.Object_ID                     " & vbNewLine & _
										" 							and c.Stereotype = 'ArchiMate_Composition'          " & vbNewLine & _
										" inner join t_object po on po.Object_ID = c.Start_Object_ID                    " & vbNewLine & _
										" 					and po.Object_Type = o.Object_Type                          " & vbNewLine & _
										" 					and isnull(po.Stereotype, '') = isnull(o.Stereotype, '')    " & vbNewLine & _
										" where o.ea_guid = '#EAGUID#'                                                  "
Class TreeNode
	Private m_SubNodes
	Private m_ParentNode
	Private m_Element
	Private m_ParentQuery
	Private m_SubNodesQuery
	
	Private Sub Class_Initialize
	  set m_SubNodes = Nothing
	  set m_ParentNode = Nothing
	  set m_Element = nothing
	  m_ParentQuery = archimateCompositionParentQuery 'default
	  m_SubNodesQuery = ""
	End Sub
	
	Public Property Get Element
		set Element = m_Element
	End Property	
	public Property Let Element(value)
		set m_Element = value
	end Property
	
	Public Property Get ParentQuery
		ParentQuery = m_ParentQuery
	End Property	
	public Property Let ParentQuery(value)
		m_ParentQuery = value
	end Property
	
	Public Property Get SubNodesQuery
		if m_SubNodesQuery = "" then
			m_SubNodesQuery = replace(me.ParentQuery, "where o.ea_guid", "where po.ea_guid")
			m_SubNodesQuery = replace(m_SubNodesQuery, "select po.Object_ID", "select o.Object_ID")
		end if
		SubNodesQuery = m_SubNodesQuery
	End Property	
	
	Public Property Get ParentNode
		if m_ParentNode is nothing then
			dim sqlGetData
			sqlGetData = replace(me.ParentQuery, "#EAGUID#", me.Element.ElementGUID)
			dim result
			set result = getElementsFromQuery(sqlGetData)
			if result.Count > 0 then
				set m_ParentNode = new TreeNode
				m_ParentNode.Element = result(0)
			end if
		end if
		set ParentNode = m_ParentNode
	End Property
	
	Public Property Get SubNodes
		if m_SubNodes is nothing then
			set m_SubNodes = CreateObject("System.Collections.ArrayList")
			dim sqlGetData
			sqlGetData = replace(me.SubNodesQuery, "#EAGUID#", me.Element.ElementGUID)
			dim result
			set result = getElementsFromQuery(sqlGetData)
			dim element as EA.Element
			for each element in result
				dim subNode
				set subNode = new TreeNode
				subNode.Element = element
				'add to list
				m_SubNodes.Add subNode
			next
		end if
		set SubNodes = m_SubNodes
	End Property
	
	public function getAllSubNodesDictionary()
		'create dictionary from the direct subNodes
		dim dictionary
		set dictionary = CreateObject("Scripting.Dictionary")
		dim subNode
		for each subNode in me.SubNodes
			'add subNode to dictionary
			dictionary.Add subNode.Element.ElementID, subNode
			'add all subNodes of the subNode to the dictionary
			dim subDictionary
			set subDictionary = subNode.getAllSubNodesDictionary()
			dim key
			for each key in subDictionary.Keys
				if not dictionary.Exists(key) then
					dictionary.Add key, subDictionary(key)
				end if
			next
		next
		'return
		set getAllSubNodesDictionary = dictionary
	end function
	
	
end Class