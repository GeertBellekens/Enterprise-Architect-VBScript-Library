'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util
'
' Script Name: 56 - Unused Business Rules
' Author: Matthias Van der Elst	
' Purpose: Lists all the business rules that are not linked to a scenario step
' Date: 2017-11-15
'

const outPutName = "56 - Unused Business Rules"
sub main
	dim usecase as EA.Element
	dim usecases 
	set usecases = CreateObject("System.Collections.ArrayList")
	dim sqlGetUsecases
	sqlGetUsecases = "select o.object_id " & _
					 "from t_object o " & _
					 "where o.Object_Type = 'usecase'"
	
	set usecases = getElementsFromQuery(sqlGetUsecases)
	
	for each usecase in usecases
		dim scenario as EA.Scenario
		for each scenario in usecase.Scenarios
			dim scenarioStep as EA.ScenarioStep
			for each scenarioStep in scenario.Steps
				
			next
		next	
	next
					 
	
end sub

main