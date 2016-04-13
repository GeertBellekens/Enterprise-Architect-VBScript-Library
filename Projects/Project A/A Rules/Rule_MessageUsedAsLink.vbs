'[path=\Projects\Project A\A Rules]
'[group=Atrias Rules]

!INC Local Scripts.EAConstants-VBScript


Class Rule_MessageUsedAsLink
'#region private attributes
	private m_Autofix
	private m_Name
	private m_ProblemStatement
	private m_Resolution
'#endregion private attributes

'#region "Constructor"
	Private Sub Class_Initialize
		m_Name = "Message used as Link"
		m_ProblemStatement = "A Message on this diagram is used as link and not as instance"
		m_Resolution = "Ctrl-Drag the message on the diagram as Instance, then run the synchronize script on it"
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
	'The Validate will validate the given item agains this rule
	public function Validate(item)
		dim validationResult
		set validationResult = new ValidationResult
		validationResult.Rule = me
		validationResult.IsValid = true
		validationResult.ValidatedItem = item
		dim invalidElementNames
		set invalidElementNames = CreateObject("System.Collections.ArrayList")
		'Rule should only be executed on a Business Process.
		if item.ObjectType = otElement then
			dim businessProcess as EA.Element
			set businessProcess = item
			if businessProcess.Stereotype = "BusinessProcess" or businessProcess.Stereotype = "Activity" then
				'get the business process diagram
				dim diagram as EA.Diagram
				set diagram = getBusinessProcessdiagram(businessProcess)
				if not diagram is nothing then
					dim diagramObject as EA.DiagramObject
					'get the BPMN activities
					for each diagramObject in diagram.DiagramObjects
						dim element as EA.Element
						set element = Repository.GetElementByID(diagramObject.ElementID)
						if (element.Stereotype = "Message" or  element.Stereotype = "FIS" _
							and element.PackageID <> diagram.PackageID) then
							'if the Message is not in the same package as the diagram then there is a problem.
							validationResult.IsValid = false
							invalidElementNames.Add element.Name
						end if
					next
				end if
			end if
		end if
		'set problem and resolution if not valid
		if validationResult.IsValid = false then
			if invalidElementNames.Count = 1 then
				validationResult.ProblemStatement = "The Message: " & invalidElementNames(0) & " is used as Link instead of as Instance"
				validationResult.Resolution = "Ctrl-Drag the message on the diagram as Instance, then run the synchronize script on it"
			else
				validationResult.ProblemStatement = "The Messages: " & Join(invalidElementNames.ToArray(),", ") & " are used as Link instead of as Instance"
				validationResult.Resolution = "Ctrl-Drag the messages on the diagram as Instance, then run the synchronize script on them"
			end if
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