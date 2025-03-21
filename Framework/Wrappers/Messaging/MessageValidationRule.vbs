'[path=\Framework\Wrappers\Messaging]
'[group=Messaging]

' Author: Geert Bellekens
' Purpose: A wrapper class for a message validation rule
' Date: 2017-03-20

Class MessageValidationRule
	'private variables
	Private m_Name
	Private m_RuleID
	private m_Reason
	private m_TestElement
	private m_Path
	
	'constructor
	Private Sub Class_Initialize
		m_Name = ""
		m_RuleID = ""
		m_Reason = ""
		set m_TestElement = nothing
		set m_Path = CreateObject("System.Collections.ArrayList")
	End Sub
	
	'public properties
	
	' Name property.
	Public Property Get Name
	  Name = m_Name
	End Property
	Public Property Let Name(value)
	  m_Name = value
	End Property
	
	' RuleId property.
	Public Property Get RuleId
	  RuleId = m_RuleId
	End Property
	Public Property Let RuleId(value)
	  m_RuleId = value
	End Property
	
	' Reason property.
	Public Property Get Reason
	  Reason = m_Reason
	End Property
	Public Property Let Reason(value)
	  m_Reason = value
	End Property	
	
	' Path property.
	Public Property Get Path
	  set Path = m_Path
	End Property
	Public Property Let Path(value)
	  set m_Path = value
	End Property
	
	'test element property
	Public Property Get TestElement
	  set TestElement = m_TestElement
	End Property
	Public Property Let TestElement(value)
	  initialiseWithTestElement value
	End Property

	
	'public operations
	public function initialiseWithTestElement(testElement)
		set m_testElement = testElement
		me.Name = Repository.GetFormatFromField("TXT",m_testElement.Notes)
		me.RuleId = m_testElement.Name
		me.Reason = getTaggedValueValue(testElement, "Error Reason")
		'get the value of the path tagged value
		dim pathString
		pathString = getTaggedValueValue(m_testElement, "Constraint Path")
		if len(pathString) > 0 then
			dim part
			for each part in Split(pathString,".")
				m_Path.Add part
			next
		end if		
	end function
	
	
end Class