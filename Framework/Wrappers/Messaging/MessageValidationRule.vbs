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
	
	'constructor
	Private Sub Class_Initialize
		m_Name = ""
		m_RuleID = ""
		m_Reason = ""
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
end Class