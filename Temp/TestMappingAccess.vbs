'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: test mapping acces
' Author: 
' Purpose: 
' Date: 
'
sub main
	dim model
	set model = getEAAddingFrameworkModel()
	dim mappingAddin
	set mappingAddin = CreateObject("EAMapping.EAMappingAddin")
	'initialize
	mappingAddin.EA_FileOpen Repository
	dim interopHelper
	set interopHelper = CreateObject("EAAddinFramework.Utilities.ScriptingInteropHelper")
	dim settings
	set settings = CreateObject("EAMapping.EAMappingSettings")
	
	dim sourceRoot
	'set sourceRoot = model.getItemFromGUID("{8D2DB628-133B-4891-AC94-B489A996C621}") 'test simple
	set sourceRoot = model.getItemFromGUID("{73A037C9-5B69-4417-9094-18A51D0633E8}") 'UNCEFACT source message
	
	dim test 
'	set test = sourceRoot.ownedAttributes
'	Session.Output sourceRoot.ownedAttributes.Count
	dim targetRootElement
	'set targetRootElement = model.getItemFromGUID("{5AE995B1-0956-49d9-8380-D8D7D5BC68BE}") 'test simple
	set targetRootElement = model.getItemFromGUID("{AA678EA4-795F-4611-B251-E286F2EE4853}") 'UNCEFACT target datamodel
	dim mappingSet
	dim parameters
	parameters = Array(sourceRoot, targetRootElement, settings)
	set mappingSet =  interopHelper.executeStaticMethod("EAAddinFramework.Mapping.MappingFactory","createMappingSet" ,parameters)
	dim mappings 
	set mappings =  interopHelper.getEnumeratedProperty(mappingSet,"mappings")
	dim mapping
	for each mapping in mappings
		Session.Output mapping.source.name
		dim mappinglogic
		for each mappingLogic in interopHelper.getEnumeratedProperty(mapping,"mappingLogics")
			dim context
			set context = mappingLogic.context
			if not context is nothing then
				Session.Output "Context: " & context.name
			end if
			Session.Output "Mapping Logic: " & mappingLogic.description
		next
	next
end sub

'gets a new instance of the EAAddinFramework and initializes it with the EA.Repository
function getEAAddingFrameworkModel()
	'Initialize the EAAddinFramework model
    dim model 
    set model = CreateObject("TSF.UmlToolingFramework.Wrappers.EA.Model")
    model.initialize(Repository)
	set getEAAddingFrameworkModel = model
end function

main