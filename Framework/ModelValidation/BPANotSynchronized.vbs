'[path=\Framework\ModelValidation]
'[group=ModelValidation]

!INC Local Scripts.EAConstants-VBScript


Class Rule_BPANotSynchronized
'#region private attributes
	private m_Autofix
	private m_Name
	private m_ProblemStatement
	private m_Resolution
'#endregion private attributes

'#region "Constructor"
	Private Sub Class_Initialize
		me.m_name = "BPA not synchronised"
		me.m_ProblemStatement = "The Business Process Activity from the library is used as Link on this diagram"
		me.m_Resolution "Execute the Synchronize script on this Business Process Activity"
		me.AutoFix = false
	end sub
'#endregion "Constructor"
	
'#region Properties
	' Autofix property.
	Public Property Get Autofix
	  Autofix = m_Autofix
	End Property
	Public Property Let Autofix(value)
	  m_Autofix = value
	End Property
	
	' Name property.
	Public Property Get Name
	  Name = m_Name
	End Property
	
	' ProblemStatement property.
	Public Property Get ProblemStatement
	  ProblemStatement = m_ProblemStatement
	End Property	
	
	' Resolution property.
	Public Property Get Resolution
	  Resolution = m_Resolution
	End Property	
'#endregion Properties
	
'#region functions
	'The Validate will validate the given item agains this rule
	public function Validate(item)
		'Rule should only be executed on a Business Process.
		if item.ObjectType = otElement then
			dim businessProcess as EA.Element
			set businessProcess = item
			if businessProcess.Stereotype = "BusinessProcess" or businessProcess.Stereotype = "Activity" then
				
			end if
		end if
	end function
	
	'the Fix function will fix the problem if possible
	' returns true if the fix succeeded and false if it wasn't able to fix the problem
	public function Fix(item, options)
	end function
'#endregion functions	
End class