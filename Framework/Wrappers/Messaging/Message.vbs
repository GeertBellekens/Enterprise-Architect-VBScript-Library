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

	'constructor
	Private Sub Class_Initialize
		m_Name = ""
		set m_RootNode = nothing
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
	  RootNode = m_RootNode
	End Property
	Public Property Let RootNode(value)
	  m_RootNode = value
	End Property
	
	public function loadMessage(eaRootNodeElement)
		'create the root node
		me.RootNode = new MessageNode
		me.RootNode.intitializeWithSource eaRootNodeElement, nothing, "1..1", nothing, nothing
	end function
	
end Class