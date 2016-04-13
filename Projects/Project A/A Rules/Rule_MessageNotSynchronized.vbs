'[path=\Projects\Project A\A Rules]
'[group=Atrias Rules]

!INC Local Scripts.EAConstants-VBScript


Class Rule_MessageNotSynchronized
'#region private attributes
	private m_Autofix
	private m_Name
	private m_ProblemStatement
	private m_Resolution
'#endregion private attributes

'#region "Constructor"
	Private Sub Class_Initialize
		m_Name = "Message not synchronised"
		m_ProblemStatement = "This Message is not synchronized with the FIS in the library"
		m_Resolution = "Execute the Synchronize script on this Message on the diagram"
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
		validationResult.ProblemStatement = me.ProblemStatement
		validationResult.Resolution = me.Resolution
		if (item.Stereotype = "Message" or  item.Stereotype = "FIS") then
		'check if this is a local instance under a business process.
			dim businessProcess as EA.Element
			set businessProcess = getOwningBusinessProcess(item)
			if not businessProcess is nothing then
				if businessProcess.PackageID = item.PackageID then
					if not item.ClassifierID > 0 then
						'if the Message does not have a classifierID then it is not synchronised
						validationResult.IsValid = false
					else 
						dim libraryMessage as EA.Element
						set libraryMessage = Repository.GetElementByID(item.ClassifierID)
						if not libraryMessage is nothing then
							if not libraryMessage.CompositeDiagram is nothing then
								if not item.CompositeDiagram is nothing then
									if item.CompositeDiagram.DiagramID <> libraryMessage.CompositeDiagram.DiagramID then
										'if both have a composite diagram then it should be the same
										validationResult.IsValid = false
									end if
								else
									'if the library message has the composite diagram then the instance should have one too
									validationResult.IsValid = false
								end if
							else
								if not item.CompositeDiagram is nothing then
									'no composite diagram for libraryMessage, so there should not be one for the instance
									validationResult.IsValid = false
								end if
							end if
						else
							'message has classifierID, but the actual message is not found
							validationResult.IsValid = false
						end if
					end if 
				end if
			end if
		end if
		'return
		set Validate = validationResult
	end function
	
	'the Fix function will fix the problem if possible
	' returns true if the fix succeeded and false if it wasn't able to fix the problem
	public function Fix(item, options)
	end function
	
	private function getOwningBusinessProcess(item)
		dim businessProcess as EA.Element
		set getOwningBusinessProcess = Nothing
		if item.ParentID > 0 then
			set businessProcess = Repository.GetElementByID(item.ParentID)
			if not businessProcess is nothing then
				if businessProcess.Stereotype = "BusinessProcess" or businessProcess.Stereotype = "Activity" then
					set getOwningBusinessProcess = businessProcess
				else
					set getOwningBusinessProcess = getOwningBusinessProcess(businessProcess)
				end if
			end if 
		end if
	end function
	
'#endregion functions	
End class