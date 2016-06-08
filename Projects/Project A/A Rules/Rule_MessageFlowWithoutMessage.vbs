'[path=\Projects\Project A\A Rules]
'[group=Atrias Rules]

!INC Local Scripts.EAConstants-VBScript


Class Rule_MessageFlowWithoutMessage
'#region private attributes
	private m_Autofix
	private m_Name
	private m_ProblemStatement
	private m_Resolution
'#endregion private attributes

'#region "Constructor"
	Private Sub Class_Initialize
		m_Name = "MessageFlow without Message"
		m_ProblemStatement = "The tagged values MessageRef on the MessageFlow is not filled in"
		m_Resolution = "Link the MessageFlow to a FIS using the tagged value MessageRef"
		AutoFix = false
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
	'The Validate will validate the given item against this rule
	public function Validate(item)
		dim item as EA.Element
		dim validationResult
		set validationResult = new ValidationResult
		validationResult.Rule = me
		validationResult.IsValid = true
		validationResult.ProblemStatement = me.ProblemStatement
		validationResult.Resolution = me.Resolution
		validationResult.ValidatedItem = item
		'Validate Messageflow
		'We do not validate connectors, se we have to validate the element sending or receiving he connector.
		'Since often the lanes/pools tend to send multiple messages it is better to us the other end (often intermediary event)
		'We have to use an SQL query to get the messageflow for this element.
		if item.Type <> "ActivityPartition" then
			'make the query to figure out if there's a message flow without reference connected to this element.
			
		end if
		'return
		set Validate = validationResult
	end function
	
	'the Fix function will fix the problem if possible
	' returns true if the fix succeeded and false if it wasn't able to fix the problem
	public function Fix(item, options)
	end function
	
	private function getBusinessProcessdiagram(businessProcess)
		dim diagram as EA.Diagram
		set getBusinessProcessdiagram = nothing
		for each diagram in businessProcess.Diagrams
			if diagram.MetaType = "BPMN2.0::Business Process" then
				set getBusinessProcessdiagram = diagram
				exit for
			end if
		next
	end function
'#endregion functions	
End class