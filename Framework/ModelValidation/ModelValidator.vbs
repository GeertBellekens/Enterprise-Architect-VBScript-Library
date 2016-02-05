'[path=\Framework\ModelValidation]
'[group=ModelValidation]


Class ModelValidator
'#region private attributes
	 private m_Rules
'#endregion private attributes

'#region "Constructor"
	Private Sub Class_Initialize
		'initialize all rules
		me.Rules = CreateObject("System.Collections.ArrayList")
		me.Rules.Add new TestRule
	end sub
'#endregion "Constructor"
	
'#region Properties
		' Rules property
	Public Property Get Rules
	  set Rules = m_Rules
	End Property
	Public Property Let Rules(value)
	  set m_Rules = value
	End Property
'#endregion Properties
	
'#region functions
	public function Validate(items, alwaysAutoFix, neverAutoFix, options, outputTabName)
		dim item, rule, validationResults,  autoFixResult, validationResult
		set validationResults = CreateObject("System.Collections.ArrayList")
		for each item in items
			Repository.WriteOutput outputTabName, "Validating item: " & getItemName(item),0
			for each rule in me.Rules
				set validationResult = rule.Validate(item)
				if (alwaysAutoFix or rule.AutoFix) _
						and not neverAutoFix and not validationResult.IsValid then
					autoFixResult = rule.Fix(item,options)
				end if
				validationResults.Add validationResult
			next
		next
		set Validate = validationResults
	end function
'#endregion functions	
End class