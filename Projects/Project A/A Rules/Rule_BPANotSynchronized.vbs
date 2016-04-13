'[path=\Projects\Project A\A Rules]
'[group=Atrias Rules]

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
		m_Name = "BPA not synchronised"
		m_ProblemStatement = "The Business Process Activity from the library is used as Link on this diagram"
		m_Resolution = "Execute the Synchronize script on this Business Process Activity"
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
						if (element.Type = "Activity" and element.Stereotype = "Activity" _
							and element.PackageID <> diagram.PackageID) then
							'if the Activity is not in the same package as the diagram then there is a problem.
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
				validationResult.ProblemStatement = "The Activity " & invalidElementNames(0) & " is using a Library Activity as link"
				validationResult.Resolution = "Synchronize the Activity to make it a local instance"
			else
				validationResult.ProblemStatement = "The activities " & Join(invalidElementNames.ToArray(),", ") & " are using Library activities as link"
				validationResult.Resolution = "Synchronize the Activities to make them local instances"
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