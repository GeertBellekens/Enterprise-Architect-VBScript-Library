'[path=\Framework\ModelValidation\Example]
'[group=ExampleModelValidation]
'
'Delete the <DISABLED> below to enable EA-Matic in your scripts 
'<DISABLED>EA-Matic
'<DISABLED>EA-Matic: http://bellekens.com/ea-matic/

option explicit 

!INC Local Scripts.EAConstants-VBScript
!INC Logging.Logger
!INC Logging.LogManager
!INC ModelValidationExample.ExampleModelValidationConstants

dim logger
set logger = new LoggerClass
logger.init "ExampleRule_<Name>"

' 
' EA_OnInitializeUserRules()
' is done in ExampleModelValidationRules_LoadRules
' A new RuleID must be created before it can be used in this file.
' ModelValidationExample.ExampleModelValidationConstants then needs to 
' defined a constant for this new rule.
'

'''''''''''''''
' Your rule should do one validation only.
' Create more rules if you need them.
' Generally your rule will only need to handle one of these events below.
' Delete all the unused event handlers.
'''''''''''''''

function EA_OnRunElementRule(RuleID, Element)
	if exampleRuleOneId <> RuleID then
		exit function
	end if

	Logger.debug "EA_OnRunElementRule called RuleId=" & RuleID & " Element.Name=" & Element.Name
	dim project as EA.Project
	set project = Repository.GetProjectInterface()
	
	'
	' Do your rule validation here.
	' Use project.PublishResult to notify any violations.
	'
	
	' The second parameter uses EnumMVErrorType values which are defined in Local Scripts.EAConstants-VBScript
	' and are mvError, mvWarning, mvInformation, mvErrorCritical.
	' The third paramter is a string for the validation message.
	' <your>RuleId is defined in ModelValidation.<your>ModelValidationConstants
	project.PublishResult exampleRuleOneId, mvInformation, "An example Info message for Element " & Element.Name
end function
