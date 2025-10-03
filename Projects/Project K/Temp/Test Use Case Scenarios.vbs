'[path=\Projects\Project K\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Test Use Case Scenarios
' Author: Geert Bellekens
' Purpose: Test the extraction of use cases scenario info
' Date: 2025-09-19
'
const outPutName = "Test Use case Scenario"

function Main ()
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	dim useCase as EA.Element
	set useCase = Repository.GetContextObject()
	if useCase is nothing then
		exit function
	end if
	'inform user
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName & "'", 0
	'do the actual work
	printScenarioInfo(useCase)
	'inform user
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName & "'", 0
		
end function

function printScenarioInfo(useCase)
	dim scenarios
	set scenarios = getScenariosForUseCase(useCase.ElementID)
	dim scenario
	for each scenario in scenarios
		Repository.WriteOutput outPutName, now() & " Scenario name: '" & scenario.Name & "'", 0
		Repository.WriteOutput outPutName, now() & " Scenario type: '" & scenario.ScenarioType & "'", 0
		dim combinedScenarioSteps
		set combinedScenarioSteps = scenario.getCombinedScenariosteps
		dim scenarioStep
		for each scenarioStep in combinedScenarioSteps
			Repository.WriteOutput outPutName, Now() & " Step Level: '" & scenarioStep.Level & "'", 0
			Repository.WriteOutput outPutName, Now() & " Step Name: '" & scenarioStep.Name & "'", 0
'			Repository.WriteOutput outPutName, Now() & " Step GUID: '" & scenarioStep.GUID & "'", 0
'			Repository.WriteOutput outPutName, Now() & " Step Uses: '" & scenarioStep.Uses & "'", 0
'			Repository.WriteOutput outPutName, Now() & " Step UsesList: '" & scenarioStep.UsesList & "'", 0
'			Repository.WriteOutput outPutName, Now() & " Step Result: '" & scenarioStep.Result & "'", 0
'			Repository.WriteOutput outPutName, Now() & " Step State: '" & scenarioStep.State & "'", 0
'			Repository.WriteOutput outPutName, Now() & " Step Trigger: '" & scenarioStep.Trigger & "'", 0
'			Repository.WriteOutput outPutName, Now() & " Step Link: '" & scenarioStep.Link & "'", 0
		next	
	next
end function

main