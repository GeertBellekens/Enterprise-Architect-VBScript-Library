'[path=\Framework\OCL]
'[group=OCL]


'Author: Geert Bellekens
'Date: 2017-11-24
'Purpose: Class containing a single OCL statement

Class OCLStatement
'#region private attributes
	private m_Context
	private m_LeftHand
	private m_NextOCLStatement
	private m_Operator
	private m_RightHand
	private m_Statement
	private m_IsValid
	
	private regexp
'#endregion private attributes

'#region "Constructor"
	Private Sub Class_Initialize
		Set regexp = CreateObject("VBScript.RegExp")
		regexp.Global = True   
		regexp.IgnoreCase = False
		set m_NextOCLStatement = nothing
	end sub
'#endregion "Constructor"
	
'#region Properties
	'Context property (EA.Element)
	Public Property Get Context
	  set Context = m_Context
	End Property
	Public Property Let Context(value)
	  set m_Context = value
	End Property
	
	'LeftHand property (string)
	Public Property Get LeftHand
	  LeftHand = m_LeftHand
	End Property
	Public Property Let LeftHand(value)
	  m_LeftHand = value
	End Property
	
	' NextOCLStatement property (OCLStatement)
	Public Property Get NextOCLStatement
	  set NextOCLStatement = m_NextOCLStatement
	End Property
	Public Property Let NextOCLStatement(value)
	  set m_NextOCLStatement = value
	End Property
	' Operator property (string) - 
	' can be "=", "->size()=", "->forAll(", "->notEmpty()"
	Public Property Get Operator
	  Operator = m_Operator
	End Property
	Public Property Let Operator(value)
	  m_Operator = replace(lcase(value), " ","")
	End Property
	
	' RightHand property (string)
	Public Property Get RightHand
	  RightHand = m_RightHand
	End Property
	Public Property Let RightHand(value)
	  m_RightHand = value
	End Property	
	' Statement property (string), contain the whole string
	Public Property Get Statement
	  Statement = m_Statement
	End Property
	Public Property Let Statement(value)
	  m_Statement = value
	  parseStatement()
	End Property
	'IsValid property (boolean)
	Public Property Get IsValid
	  IsValid = m_IsValid
	End Property
	Public Property Let IsValid(value)
	  m_IsValid = value
	End Property
'#endregion Properties
	
'#region functions
	'Show this resultset in the model search window
	public function parseStatement()
		'remove leading whitespace
		regExp.Pattern = "^\s*"
		m_Statement = regExp.Replace(me.Statement, "")
		'remove trailing whitespace
		regExp.Pattern = "\s*$"
		m_Statement = regExp.Replace(me.Statement, "")
		dim matches
		'group into left(1), operator(2) and right(3)
		regExp.Pattern = "(^.*?)(->size\(\)=?|->notEmpty\(\)|->forAll ?\(|=)(?=[\s\S]+$)?([\s\S]+$)?"
		set matches = regExp.Execute(me.Statement)
		dim match
		if matches.Count = 1 then
			set match = matches(0)
			if match.SubMatches.Count >=2 then
				'create new OCL statements
				me.LeftHand = match.SubMatches(0)
				me.Operator = match.SubMatches(1)
				if match.SubMatches.Count = 3 then
					me.RightHand = match.SubMatches(2)
					me.IsValid = true 'always valid if 
					me.ParseSubStatements()				
				else
					if me.Operator = "->notEmpty()" then 'only case in which there are only two matches
						me.IsValid = true
					else
						me.IsValid = false
					end if
				end if

			else
				'Indicate not valid
				me.IsValid = false
			end if
		else
			'there should only be one match
			me.IsValid = false
		end if
		
	end function
	
	function ParseSubStatements()
		'check if the statement exists of multiple statements (starting with self.)
		regExp.Pattern = "(\S*)[\s]*?or[ \s]*(self\.[\s\S]*)$"
		dim matches
		set matches = regExp.Execute(me.RightHand)
		if matches.Count = 1 then
			set match = matches(0)
			if match.SubMatches.Count =2 then
				'get the limited right part
				me.RightHand = match.SubMatches(0)
				'create next statement
				me.NextOCLStatement = new OCLStatement
				me.NextOCLStatement.Context = me.Context
				me.NextOCLStatement.Statement = match.SubMatches(1)
			end if
		else
			'check if there are multiple statements for a value (used in Forall ( x = y or w = z))
			regExp.Pattern = "(\S*) ?= ?(\S*)[\s]*?or[ \s]*([\s\S]*\)\s*)$"
			set matches = regExp.Execute(me.RightHand)
			if matches.Count = 1 then
				set match = matches(0)
				if match.SubMatches.Count =3 then
					'create next statement
					me.NextOCLStatement = new OCLStatement
					me.NextOCLStatement.Context = me.Context
					me.NextOCLStatement.Statement = me.LeftHand & me.Operator & match.SubMatches(2)
					'then reset this statement to simple "=" statement
					me.LeftHand = me.LeftHand & "." & match.SubMatches(0)
					me.Operator = "="
					me.RightHand = match.SubMatches(1)
				end if
			else
				'check if partial statement between brackets (used in ForAll (x=y))
				if lcase(left(me.Operator,8)) = "->forall" then
					regExp.Pattern = "^\s*?(\S*) ?= ?(\S*)\s*?\)\s*?$"
					set matches = regExp.Execute(me.RightHand)
					if matches.Count = 1 then
						set match = matches(0)
						if match.SubMatches.Count =2 then
							me.LeftHand = me.LeftHand & "." & match.SubMatches(0)
							me.Operator = "="
							me.RightHand = match.SubMatches(1)
						end if
					end if
				end if
			end if
		end if
	end function
'#endregion functions	
end class