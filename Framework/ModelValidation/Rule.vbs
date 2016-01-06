'[path=\Framework\ModelValidation]
'[group=ModelValidation]

!INC Local Scripts.EAConstants-VBScript


Class Rule
'#region private attributes
	private m_Autofix
	private m_Name
'#endregion private attributes

'#region "Constructor"
	Private Sub Class_Initialize
		m_name = "AbstractRule"
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
'#endregion Properties
	
'#region functions
	'The Validate will validate the given item agains this rule
	'It returns a ValidationResult
	public function Validate(item)
		
	end function
	
	'the Fix function will fix the problem if possible
	' returns true if the fix succeeded and false if it wasn't able to fix the problem
	public function Fix(item, options)
	end function
'#endregion functions	
End class