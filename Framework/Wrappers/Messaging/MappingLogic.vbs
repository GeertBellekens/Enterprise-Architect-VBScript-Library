'[path=\Framework\Wrappers\Messaging]
'[group=Messaging]

!INC Utils.Include

' Author: Geert Bellekens
' Purpose: A wrapper class for a mapping logic
' Date: 2019-03-08

Class MappingLogic
	'private variables
	Private m_Context
	Private m_Description

	
	'constructor
	Private Sub Class_Initialize
		set m_Context = nothing
		m_Description = ""
	End Sub
	
	'public properties
	
	' Context property.
	Public Property Get Context
	  set Context = m_Context
	End Property
	Public Property Let Context(value)
	 set m_Context = value
	End Property
	
	' Description property.
	Public Property Get Description
	  Description = m_Description
	End Property
	Public Property Let Description(value)
		m_Description = value
	End Property
	
end Class