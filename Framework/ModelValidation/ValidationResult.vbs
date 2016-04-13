'[path=\Framework\ModelValidation]
'[group=ModelValidation]

Class ValidationResult
'#region private attributes
	private m_ValidatedItem
	private m_IsValid
	private m_Rule
	private m_Headers
	private m_ProblemStatement
	private m_Resolution
'#endregion private attributes

'#region "Constructor"
	Private Sub Class_Initialize
		me.IsValid = false
		me.ValidatedItem = nothing
		me.Rule = nothing
		set m_Headers =  CreateObject("System.Collections.ArrayList")
		m_Headers.Add "CLASSGUID"
		m_Headers.Add "CLASSTYPE"
		m_Headers.Add "Name"
		m_Headers.Add "Rule"
		m_Headers.Add "IsValid"
		m_Headers.Add "Problem"
		m_Headers.Add "Resolution"
		m_Headers.Add "Fully Qualified Name"
	end sub
'#endregion "Constructor"
	
'#region Properties
	' ValidatedItem property
	Public Property Get ValidatedItem
	  set ValidatedItem = m_ValidatedItem
	End Property
	Public Property Let ValidatedItem(value)
	  set m_ValidatedItem = value
	End Property
	
	' IsValid property
	Public Property Get IsValid
	  IsValid = m_IsValid
	End Property
	Public Property Let IsValid(value)
	  m_IsValid = value
	End Property
	
	' Rule property
	Public Property Get Rule
	  set Rule = m_Rule
	End Property
	Public Property Let Rule(value)
	  set m_Rule = value
	End Property
	
	'Headers property
	Public Property Get Headers
		Set Headers = m_Headers
	end Property
	
	' ProblemStatement property
	Public Property Get ProblemStatement
	  if len(m_ProblemStatement) = 0 then
		ProblemStatement = me.Rule.ProblemStatement
	  else
		ProblemStatement = m_ProblemStatement
	  end if
	End Property
	Public Property Let ProblemStatement(value)
	  m_ProblemStatement = value
	End Property
	
	' Resolution property
	Public Property Get Resolution
	  if len(m_Resolution) = 0 then
		Resolution = me.Rule.Resolution
	  else
		Resolution = m_Resolution
	  end if
	End Property
	Public Property Let Resolution(value)
	  m_Resolution = value
	End Property

'#endregion Properties
	
'#region functions
	'Returns an ArrayList containing the headers of the results
	public function getResultFields()
		dim resultFields
		set resultFields = CreateObject("System.Collections.ArrayList")
		resultFields.Add me.ValidatedItem.ElementGUID
		resultFields.Add me.ValidatedItem.Type
		resultFields.Add me.ValidatedItem.Name
		resultFields.Add me.Rule.Name
		if me.IsValid then
			resultFields.Add "True"
		else
			resultFields.Add "False"
		end if
		resultFields.Add me.ProblemStatement
		resultFields.Add me.Resolution
		resultFields.Add getFullyQualifiedName(me.ValidatedItem)
'		resultFields.Add "Qualified Name takes too long"
		set getResultFields = resultFields
	end function
	

'#endregion functions	
End class