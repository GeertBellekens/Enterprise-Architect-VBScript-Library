'[path=\Framework\ModelValidation]
'[group=ModelValidation]
'
'Delete the <DISABLED> below to enable EA-Matic in your scripts 
'<DISABLED>EA-Matic
'<DISABLED>EA-Matic: http://bellekens.com/ea-matic/

option explicit 

!INC Local Scripts.EAConstants-VBScript
!INC Logging.Logger
!INC Logging.LogManager
!INC ModelValidation.<your>ModelValidationConstants

dim logger
set logger = new LoggerClass
logger.init "<your>ModelValidationRule_<Name>"

' 
' EA_OnInitializeUserRules()
' is done in <your>ModelValidationRules_LoadRules
' A new RuleID must be created before it can be used in this file.
' ModelValidation.<your>ModelValidationConstants then needs to 
' defined a constant for this new rule.
'

'''''''''''''''
' Your rule should do one validation only.
' Create more rules if you need them.
' Generally your rule will only need to handle one of these events below.
' Delete all the unused event handlers.
'''''''''''''''

function EA_OnRunElementRule(RuleID, Element)
	if <your>RuleId <> RuleID then
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
	project.PublishResult <your>RuleId, mvInformation, "An example Info message for Element " & Element.Name
end function

function EA_OnRunPackageRule(RuleID, PackageID)
	Logger.debug "EA_OnRunPackageRule called RuleId=" & RuleID & " PackageID=" & CStr(PackageID)
end function

function EA_OnRunDiagramRule(RuleID, DiagramID)
	Logger.debug "EA_OnRunDiagramRule called RuleId=" & RuleID & " DiagramID=" & CStr(DiagramID)
end function

function EA_OnRunConnectorRule(RuleID, ConnectorID)
	Logger.debug "EA_OnRunConnectorRule called RuleId="' & RuleID & " ConnectorID=" & ConnectorID
end function

function EA_OnRunAttributeRule(RuleID, AttributeGUID, ObjectID)
	Logger.debug "EA_OnRunAttributeRule called RuleId=" & RuleID & " AtttributeGUID=" & AttributeGUID & " ObjectID=" & CStr(ObjectID)
end function

function EA_OnRunMethodRule(RuleID, MethodGUID, ObjectID)
	Logger.debug "EA_OnRunMethodRule called RuleId=" & RuleID & " MethodGUID=" & MethodGUID & " ObjectID=" & CStr(ObjectID)
end function

function EA_OnRunParameterRule(RuleID, ParameterGUID, MethodGUID, ObjectID)
	Logger.debug "EA_OnRunParameterRule called RuleId=" & RuleID & " ParameterGUID=" & ParameterGUID & " MethodGUID=" & MethodGUID & " ObjectID=" & CStr(ObjectID)
end function
