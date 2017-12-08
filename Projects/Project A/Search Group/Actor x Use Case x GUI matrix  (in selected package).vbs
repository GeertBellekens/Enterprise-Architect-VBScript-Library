'[path=\Projects\Project A\Search Group]
'[group=Search Group]

' Script Name: Use Case UI Matrix (in selected package)
' Author: Geert Bellekens
' Purpose: shows the Actor x Use Case x GUI matrix in the search results
' Date: 2017-02-15

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

const outputTabName = "Actor x UseCase x GUI Matrix"

sub main
	Repository.CreateOutputTab outputTabName
	Repository.ClearOutput outputTabName
	Repository.EnsureOutputVisible outputTabName
	'get the current package ID string
	dim currentPackageIDString
	currentPackageIDString = getCurrentPackageTreeIDString()
	'get all use cases (in the selected package)
	dim allUseCases
	set allUseCases = getAllUseCases(currentPackageIDString)
	Repository.WriteOutput outputTabName, now() & " : Number of use cases found: " & allUseCases.Count ,0
	'get all included use cases for each use case
	dim includedUseCasesWithExecutionRights
	set includedUseCasesWithExecutionRights = getAllIncludesWithExecutionRights(allUseCases,currentPackageIDString)
	'show the output
	showOutput includedUseCasesWithExecutionRights
end sub 


function getAllIncludesWithExecutionRights(allUseCases,currentPackageIDString)
	dim useCase as EA.Element
	dim includesUseCasesforUseCase
	set includesUseCasesforUseCase = CreateObject("Scripting.Dictionary")
	for each useCase in allUseCases 
		dim includedUseCases
		set includedUseCases = nothing
		Repository.WriteOutput outputTabName, now() & " Processing use case: " & useCase.Name , useCase.ElementID
		set includedUseCases = getIncludedUseCases(useCase, includedUseCases,currentPackageIDString)
		includesUseCasesforUseCase.Add useCase, includedUseCases
	next
	dim outputRows
	set outputRows = CreateObject("System.Collections.ArrayList")
	'create output format, and arraylist of arraylists that list actor, use case and GUI element
	for each useCase in includesUseCasesforUseCase.Keys
		'get the actors for this use case
		dim actors
		set actors = getActors(useCase)
		dim includedUseCase
		'add rows for this use case
		addOutputRows usecase, actors, true , outputRows
		for each includedUseCase in includesUseCasesforUseCase(useCase)
			'add rows for included use case
			addOutputRows includedUseCase, actors, false , outputRows
		next
	next
	'return output
	set getAllIncludesWithExecutionRights = outputRows
end function

function addOutputRows(usecase, actors, directIndicator,outputRows)
	Repository.WriteOutput outputTabName, now() & " Creating output for use case: " & usecase.Name , usecase.ElementID
	'get the user interfaces
	dim userInterfaces 
	set userInterfaces = getUserInterfaces(usecase)
	'get packages
	'get package0
	dim package0 as EA.Package
	set package0 = Repository.GetPackageByID(useCase.PackageID)
	'get package1
	dim package1 as EA.Package
	if package0.ParentID > 0 then
		set package1 = Repository.GetPackageByID(package0.ParentID)
	end if
	'get package2
	dim package2 as EA.Package
	if not package1 is nothing and package1.ParentID > 0 then
		set package2 = Repository.GetPackageByID(package1.ParentID)
	else
		set package2 = nothing
	end if
	'get package2
	dim package3 as EA.Package
	if not package2 is nothing and package2.ParentID > 0 then
		set package3 = Repository.GetPackageByID(package2.ParentID)
	else
		set package3 = nothing
	end if
	dim row
	dim actor as EA.Element
	dim userInterface as EA.Element
	for each actor in actors.Keys
		dim actorType
		actorType = getActorType(actor)
		dim actorInheritance
		actorInheritance = actors.Item(Actor)
		'add rows for each actor
		if userInterfaces.Count = 0 then
			set row = getOutputRow(usecase, actor, actorType, actorInheritance, directIndicator, nothing, package0, package1, package2, package3)
			outputRows.Add row
		end if
		'add rows for each user interface
		for each userInterface in userInterfaces
			set row = getOutputRow(usecase, actor, actorType, actorInheritance, directIndicator, userInterface, package0, package1, package2, package3)
			outputRows.Add row
		next
	next
end function

function getActorType(actor)
	dim actorPackage as EA.Package
	set actorPackage = Repository.GetPackageByID(actor.PackageID)
	if instr(actorPackage.Name, "Human") > 0 then
		getActorType = "Human"
	else
		getActorType = "System"
	end if
end function

function getOutputRow(usecase, actor, actorType,actorInheritance, directIndicator, userInterface, package0, package1, package2, package3)
	'create outputrow
	dim outputRow
	set outputRow = CreateObject("System.Collections.ArrayList")
	'add data
	outputRow.Add useCase.ElementGUID
	outputRow.Add useCase.Type
	outputRow.Add actor.Name
	outputRow.Add actorType
	outputRow.Add actorInheritance
	outputRow.Add useCase.Name
	if directIndicator then
		outputRow.Add "Direct"
	else
		outputRow.Add "Indirect"
	end if
	if not userInterface is nothing then
		outputRow.Add userInterface.Name
	else
		outputRow.Add ""
	end if
	if not package0 is nothing then
		outputRow.Add package0.Name
	else
		outputRow.Add ""
	end if
	if not package1 is nothing then
		outputRow.Add package1.Name
	else
		outputRow.Add ""
	end if
	if not package2 is nothing then
		outputRow.Add package2.Name
	else
		outputRow.Add ""
	end if
	if not package3 is nothing then
		outputRow.Add package3.Name
	else
		outputRow.Add ""
	end if
	'return row
	set getOutputRow = outputRow
end function

function showOutput(outputRows)
	
	'get the headers for the output
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.Add "CLASSGUID"
	headers.Add "CLASSTYPE"
	headers.Add "Actor"
	headers.Add "Actor Type"
	headers.Add "Inheritance"
	headers.Add "Use Case"
	headers.Add "Use Case Link Type"
	headers.Add "User Interface"
	headers.Add "Package_level1 "
	headers.Add "Package_level2"
	headers.Add "Package_level3"
	headers.Add "Package_level4"
	'create the output object
	dim searchOutput
	set searchOutput = new SearchResults
	searchOutput.Name = "Use Case UI Matrix"
	searchOutput.Fields = headers
	'put the contents in the output
	dim row
	for each row in outputRows
		'add row the the output
		searchOutput.Results.Add row
	next
	'show the output
	searchOutput.Show
end function

function getUserInterfaces(usecase)
	'find user interfaces
	dim getUserInterfacesSQL
	getUserInterfacesSQL =  "select ui.Object_ID                                                                 " & _
							" from  t_object ac                                                                  " & _
							" inner join t_object step on step.ParentID = ac.Object_ID                           " & _
							" 							and step.Object_Type = 'Action'                          " & _
							" inner join t_connector step_ui on step_ui.Start_Object_ID = step.Object_ID         " & _
							" inner join t_object ui on step_ui.End_Object_ID = ui.Object_ID                     " & _
							" 							and ui.Stereotype = 'ArchiMate_ApplicationInterface'     " & _
							" where ac.Object_Type = 'Activity'                                                  " & _
							" and ac.ParentID = " & usecase.ElementID
	'return the userinterface
	set getUserInterfaces = getElementsFromQuery(getUserInterfacesSQL)
end function

function getActors(useCase)
	'first get the direct actors
	dim getDirectActorsSQL
	getDirectActorsSQL = 	"select act.Object_ID from t_object act                                                 " & _
							" inner join t_connector act_uc on act_uc.Start_Object_ID = act.Object_ID               " & _
							" 								and act_uc.Connector_Type in ('Association','UseCase')  " & _
							" 								and act_uc.End_Object_ID = " & useCase.ElementID & " 	" & _
							" where act.Object_Type = 'Actor'	                                                 	"
	dim actors
	set actors = getElementsFromQuery(getDirectActorsSQL)
	dim allActors
	set allActors = CreateObject("Scripting.Dictionary")
	'then get the specialized actors for each actor
	dim actor as EA.Element
	for each actor in actors
		if not allActors.Exists(actor) then
			'add the actor itself
			allActors.Add actor, "primary"
			set allActors = getSpecializedActors(actor, allActors)
		end if
	next
	'return the actors
	set getActors = allActors
end function

function getSpecializedActors(actor, allActors)
	dim allActorsIDString
	allActorsIDString = makeIDString(allActors)
	dim getChildActorsSQL
	getChildActorsSQL = "select act.Object_ID from  t_object act                                          " & _
						" inner join t_connector gen on gen.Start_Object_ID = act.Object_ID               " & _
						" 					and gen.Connector_Type in ('Generalization','Generalisation') " & _
						" 					and gen.End_Object_ID = " & actor.ElementID & "               " & _
						" where act.Object_Type = 'Actor'												  " & _
						" and act.Object_ID not in (" & allActorsIDString & ")							  "
	dim childActors
	set childActors = getElementsFromQuery(getChildActorsSQL)
	'go level deeper
	dim childActor
	for each childActor in childActors
		'add it to the list of all actors
		if not allActors.Exists(childActor) then
			allActors.Add childActor, "inherited"
			set allActors = getSpecializedActors(childActor, allActors)
		end if
	next
	'return the list of actors
	set getSpecializedActors = allActors
end function

function getIncludedUseCases(useCase, includedUseCases,currentPackageIDString)
	dim sqlGetIncludes
	dim includedUseCasesIdString
	dim directIncludes
	if not includedUseCases is nothing then
		includedUseCasesIdString = makeIDString(includedUseCases)
	else
		includedUseCasesIdString = "0"
	end if
	sqlGetIncludes = "select uc.Object_ID from t_object uc                                                " & _
					" inner join t_connector c on c.End_Object_ID = uc.Object_ID                          " & _
					" where c.Connector_Type in ('Association','UseCase')                                 " & _
					" and uc.Object_Type = 'UseCase'                                                      " & _
					" and c.Start_Object_ID = " & useCase.ElementID & "                                   " & _
					" and uc.Object_ID not in (" &  includedUseCasesIdString & ")                         " & _
					" and uc.Package_ID in (" &  currentPackageIDString & ")                              " & _
					" and not exists                                                                      " & _
					" 	(                                                                                 " & _
					" 		select act.Object_ID from t_object act                                        " & _
					" 		inner join t_connector act_c on act_c.Start_Object_ID = act.Object_ID         " & _
					" 		inner join t_package act_p on act_p.Package_ID = act.Package_ID               " & _
					" 		where act_c.Connector_Type in ('Association','UseCase')                       " & _
					" 		and act_c.End_Object_ID = uc.Object_ID                                        " & _
					" 		and act.Object_Type = 'Actor'                                                 " & _
					" 		and act_p.Name like '%Human%'                                                 " & _
					" 	)"	
	'get the new includes
	set directIncludes = getElementsFromQuery(sqlGetIncludes)
	dim includedUseCase as EA.Element
	'if the includes usecases have not been defined yet initialize them with the newincludes
	if includedUseCases is nothing then
		set includedUseCases = CreateObject("System.Collections.ArrayList")
	end if
	'add the direct included use cases
	includedUseCases.AddRange(directIncludes)

	'go one level deeper
	'loop the direct use cases
	for each includedUseCase in directIncludes
		'get their included use cases
		set includedUseCases = getIncludedUseCases(includedUseCase,includedUseCases,currentPackageIDString )
	next
	'return the included use cases
	set getIncludedUseCases = includedUseCases
end function

function getAllUseCases(currentPackageIDString)
	dim SQLGetAllUseCases
	SQLGetAllUseCases = "select uc.Object_ID " & _
						" from t_object uc " & _
						" where uc.Object_Type = 'UseCase' " & _
						" and uc.Package_ID IN (" & currentPackageIDString & ")"
	set getAllUseCases = getElementsFromQuery(SQLGetAllUseCases)
end function



main