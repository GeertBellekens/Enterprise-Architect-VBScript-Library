'[path=\Framework\Wrappers]
'[group=Wrappers]



!INC Utils.Include

'
' Script Name: UsecaseScenarioStep
' Author: Geert Bellekens
' Purpose: Represents a step in a Use case Scenario
' Date: 2025-09-19
'

Class UsecaseScenarioStep

	Private m_Name
	Private m_GUID
	Private m_Level
	Private m_Uses
	Private m_UsesList
	Private m_Result
	Private m_State
	Private m_Trigger
	Private m_Link

	Private Sub Class_Initialize()
		m_Name      = ""
		m_GUID      = ""
		m_Level     = 0
		m_Uses      = ""
		m_UsesList  = ""
		m_Result    = ""
		m_State     = ""
		m_Trigger   = ""
		m_Link      = ""
	End Sub

	' Name property.
	Public Property Get Name
		Name = m_Name
	End Property
	Public Property Let Name(value)
		m_Name = value
	End Property

	' GUID property.
	Public Property Get GUID
		GUID = m_GUID
	End Property
	Public Property Let GUID(value)
		m_GUID = value
	End Property

	' Level property.
	Public Property Get Level
		Level = m_Level
	End Property
	Public Property Let Level(value)
		m_Level = value
	End Property

	' Uses property.
	Public Property Get Uses
		Uses = m_Uses
	End Property
	Public Property Let Uses(value)
		m_Uses = value
	End Property

	' UsesList property.
	Public Property Get UsesList
		UsesList = m_UsesList
	End Property
	Public Property Let UsesList(value)
		m_UsesList = value
	End Property

	' Result property.
	Public Property Get Result
		Result = m_Result
	End Property
	Public Property Let Result(value)
		m_Result = value
	End Property

	' State property.
	Public Property Get State
		State = m_State
	End Property
	Public Property Let State(value)
		m_State = value
	End Property

	' Trigger property.
	Public Property Get Trigger
		Trigger = m_Trigger
	End Property
	Public Property Let Trigger(value)
		m_Trigger = value
	End Property

	' Link property.
	Public Property Get Link
		Link = m_Link
	End Property
	Public Property Let Link(value)
		m_Link = value
	End Property
	
	public function initialize(stepNode)
		    Me.Name      = stepNode.getAttribute("name")
			Me.GUID      = stepNode.getAttribute("guid")
			Me.Level     = Cint(stepNode.getAttribute("level"))
			Me.Uses      = stepNode.getAttribute("uses")
			Me.UsesList  = stepNode.getAttribute("useslist")
			Me.Result    = stepNode.getAttribute("result")
			Me.State     = stepNode.getAttribute("state")
			Me.Trigger   = stepNode.getAttribute("trigger")
			Me.Link      = stepNode.getAttribute("link")
	end function

	
	
end class