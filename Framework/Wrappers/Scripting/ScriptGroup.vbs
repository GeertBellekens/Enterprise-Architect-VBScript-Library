'[path=\Framework\Wrappers\Scripting]
'[group=Wrappers]
option explicit

!INC Util.Include

'Author: Geert Bellekens
'Date: 2015-12-07

Class ScriptGroup 
	Private m_Id
	Private m_Name
	Private m_GroupType
	Private m_Scripts
	
	Private Sub Class_Initialize
	  m_Id = ""
	  m_Name = ""
	  m_GroupType = ""
	  set m_Scripts = CreateObject("System.Collections.ArrayList")
	End Sub

	' Id property.
	Public Property Get Id
	  Id = m_Id
	End Property
	Public Property Let Id(value)
	  m_Id = value
	End Property	
	
	' Name property.
	Public Property Get Name
	  Name = m_Name
	End Property
	Public Property Let Name(value)
	  m_Name = value
	End Property

	' GroupType property.
	Public Property Get GroupType
	  GroupType = m_GroupType
	End Property
	Public Property Let GroupType(value)
	  m_GroupType = value
	End Property
	
	' Scripts property.
	Public Property Get Scripts
	  set Scripts = m_Scripts
	End Property
	
	'the notes contain soemthing like <Group Type="NORMAL" Notes=""/>
	'so the group type is the second part when splitted by double quotes
	private function getGroupTypeFromNotes(notes)
		dim parts
		parts = split(notes,"""")
		getGroupTypeFromNotes = parts(1)
	end function
	
	'sets the GroupType based on the given notes
	public sub setGroupTypeFromNotes(notes)
		GroupType = getGroupTypeFromNotes(notes)
	end sub

end Class