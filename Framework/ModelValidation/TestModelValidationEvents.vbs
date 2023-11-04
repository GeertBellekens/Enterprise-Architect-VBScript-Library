'[path=\Framework\ModelValidation]
'[group=ModelValidation]
'
'Delete the <DISABLED> below to enable EA-Matic for the Test Scripts
'<DISABLED>EA-Matic
'<DISABLED>EA-Matic: http://bellekens.com/ea-matic/

option explicit 

!INC Local Scripts.EAConstants-VBScript
!INC Logging.Logger
!INC Logging.LogManager
!INC ModelValidation.TestModelValidationConstants

dim logger
set logger = new LoggerClass
logger.init "TestModelValidationEvents"

' 
' EA_OnInitializeUserRules()
' is done in TestModelValidationRules_LoadRules
'

'''''''''''''''
' Model Validation Events
'
function EA_OnStartValidation(Args)
	dim ruleCategoriesAsString, i, val
	Logger.debug "EA_OnStartValidation VarType(Args)=" & VarType(Args)
	Logger.debug "EA_OnStartValidation IsArray(Args)=" & IsArray(Args)
	Logger.debug "EA_OnStartValidation LBound(Args)=" & LBound(Args)
	Logger.debug "EA_OnStartValidation UBound(Args)=" & UBound(Args)
	ruleCategoriesAsString = "["
	for i = LBound(Args) to UBound(Args)
		Logger.debug "EA_OnStartValidation i=" & i	
		' This is just not an array and you can't index it.
		' Keep getting Variable uses an Automation type not supported in VBScript
		ruleCategoriesAsString = ruleCategoriesAsString
		if i <> UBound(Args) then
			ruleCategoriesAsString = ruleCategoriesAsString & ", "
		end if
	next
	ruleCategoriesAsString = ruleCategoriesAsString & "]"
	Logger.debug "EA_OnStartValidation called args=" & ruleCategoriesAsString	
end function

function EA_OnEndValidation(Args)
	Logger.debug "EA_OnEndValidation called args=" ' & ruleCategoriesAsString
end function

function EA_OnRunElementRule(RuleID, Element)
	Logger.debug "EA_OnRunElementRule called RuleId=" & RuleID & " Element.Name=" & Element.Name
	Logger.debug "EA_OnRunElementRule testRuleOneId=" & testRuleOneId
	dim project as EA.Project
	set project = Repository.GetProjectInterface()
	project.PublishResult testRuleOneId, mvInformation, "An example Info message for Element " & Element.Name
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
