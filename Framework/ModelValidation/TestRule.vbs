'[path=\Framework\ModelValidation]
'[group=ModelValidation]


Class TestRule
'#region private attributes
 private m_Autofix
 private m_Name
 private m_ProblemStatement
 private m_Resolution
'#endregion private attributes

'#region "Constructor"
 Private Sub Class_Initialize
  m_name = "Test Rule"
  m_ProblemStatement = "There is a problem with this element"
  m_Resolution = "Fix it dammit!"
  me.AutoFix = true
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
 'It the validation result
 public function Validate(item)
  dim validationResult
  set validationResult = new ValidationResult
  validationResult.Rule = me
  validationResult.IsValid = false
  validationResult.ValidatedItem = item
  validationResult.ProblemStatement = "the element " & getItemName(item) & " is not valid"
  validationResult.Resolution = "fix " & getItemName(item) & " and fast!"
  set Validate = validationResult
 end function
 
 'the Fix function will fix the problem if possible
 ' returns true if the fix succeeded and false if it wasn't able to fix the problem
 public function Fix(item, options)
  'msgbox "Fixed it!"
 end function
'#endregion functions 
End class