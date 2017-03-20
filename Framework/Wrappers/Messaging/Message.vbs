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

	'constructor
	Private Sub Class_Initialize
		m_Name = ""
		set m_RootNode = nothing
		m_MessageDepth = 0
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
	
	public function loadMessage(eaRootNodeElement)
		'create the root node
		me.RootNode = new MessageNode
		me.RootNode.intitializeWithSource eaRootNodeElement, nothing, "1..1", nothing, nothing
	end function
	
	'create an arraylist of arraylists with the details of this message
	public function createOuput()
		dim outputList
		'create empty list for current path
		dim currentPath
		set currentPath = CreateObject("System.Collections.ArrayList")
		'start with the rootnode
		set outputList = me.RootNode.getOuput(currentPath,me.MessageDepth)
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
		'first one is always "Message"
		headers.Add("Message")
		'add the levels
		dim i
		for i = 1 to me.MessageDepth step +1
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
	
end Class