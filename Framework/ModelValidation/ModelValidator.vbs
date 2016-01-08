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
	public function Validate(items, alwaysAutoFix, neverAutoFix, options)
		dim item, rule, validationResults, isValid, autoFixResult, validationResult
		set validationResults = CreateObject("System.Collections.ArrayList")
		for each item in items
			Session.Output "Validating item " & item.Name
			for each rule in me.Rules
				isValid = rule.Validate(item)
				if (alwaysAutoFix or rule.AutoFix) _
						and not neverAutoFix then
					autoFixResult = rule.Fix(item,options)
				end if
				set validationResult = new ValidationResult
				validationResult.Rule = rule
				validationResult.IsValid = isValid
				validationResult.ValidatedItem = item
				validationResults.Add validationResult
			next
		next
		set Validate = validationResults
	end function
'#endregion functions	
End class