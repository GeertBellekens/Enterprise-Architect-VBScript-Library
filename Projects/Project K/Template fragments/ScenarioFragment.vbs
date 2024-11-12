'[path=\Projects\Project K\Template fragments]
'[group=Template fragments]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
' Script Name: ScenarioFragment
' Author: Geert Bellekens
' Purpose: Return an XML string containing the details of the given scenarioName
' Date: 2023-03-03
'
const outPutName = "UCScenarioFragment"

function test
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'test with a use case
	msgbox documentUseCaseScenarios(3009)
end function
'call test
'test

function documentUseCaseScenarios(objectID)
'	'create output tab
'	Repository.CreateOutputTab outPutName
'	Repository.ClearOutput outPutName
'	Repository.EnsureOutputVisible outPutName
'	Repository.WriteOutput outPutName, now() & " Generating UC with ID" & objectID, 0
	dim docgen as EA.DocumentGenerator
	set docgen = Repository.CreateDocumentGenerator()
	docgen.NewDocument "UC_ScenarioSteps"
	dim scenarios
	set scenarios = getScenariosForUseCase(objectID)
	dim scenario
	for each scenario in scenarios
		dim scenarioXmlString
		scenarioXmlString = getScenarioXML(scenario)
		docgen.DocumentCustomData scenarioXmlString, 1, "UC_Scenario"
		dim scenarioStepsXmlString
		scenarioStepsXmlString = getScenarioStepsXMLString(scenario)
		docgen.DocumentCustomData scenarioStepsXmlString, 1, "UC_ScenarioSteps"
	next
	documentUseCaseScenarios = docgen.GetDocumentAsRTF()
	'docgen.SaveDocument "c:\temp\docgenTest.docx", 3
end function


function getScenarioXML(scenario)
	dim xmlDOM 
	set  xmlDOM = CreateObject( "Microsoft.XMLDOM" )
	
	xmlDOM.validateOnParse = false
	xmlDOM.async = false
	 
	dim node 
	set node = xmlDOM.createProcessingInstruction( "xml", "version='1.0'")
    xmlDOM.appendChild node
'
	dim xmlRoot 
	set xmlRoot = xmlDOM.createElement( "EADATA" )
	xmlDOM.appendChild xmlRoot

	dim xmlDataSet
	set xmlDataSet = xmlDOM.createElement( "Dataset_0" )
	xmlRoot.appendChild xmlDataSet
	 
	dim xmlData 
	set xmlData = xmlDOM.createElement( "Data" )
	xmlDataSet.appendChild xmlData
	
	dim xmlRow
	set xmlRow = xmlDOM.createElement( "Row" )
	xmlData.appendChild xmlRow
	
	'ScenarioType
	dim xmlScenarioType
	set xmlScenarioType = xmlDOM.createElement("ScenarioType")	
	xmlScenarioType.text = scenario.ScenarioType
	xmlRow.appendChild xmlScenarioType
	
	'Name
	dim xmlName
	set xmlName = xmlDOM.createElement("name")	
	xmlName.text = scenario.Name
	xmlRow.appendChild xmlName
	
	'Entry
	dim xmlEntry
	set xmlEntry = xmlDOM.createElement("Entry")	
	xmlEntry.text = scenario.Entry
	xmlRow.appendChild xmlEntry
	
	'Join
	dim xmlJoin
	set xmlJoin = xmlDOM.createElement("Join")	
	xmlJoin.text = scenario.Join
	xmlRow.appendChild xmlJoin
	
	'Notes
	dim formattedAttr 
	set formattedAttr = xmlDOM.createAttribute("formatted")
	formattedAttr.nodeValue="1"
	dim xmlNotes
	set xmlNotes = xmlDOM.createElement("Notes")	
	xmlNotes.text = scenario.Notes
	xmlNotes.setAttributeNode(formattedAttr)
	xmlRow.appendChild xmlNotes
	
	'return
	getScenarioXML = xmlDOM.xml
end function

function getScenarioStepsXMLString(scenario)

	dim xmlDOM 
	set  xmlDOM = CreateObject( "Microsoft.XMLDOM" )
	
	xmlDOM.validateOnParse = false
	xmlDOM.async = false
	 
	dim node 
	set node = xmlDOM.createProcessingInstruction( "xml", "version='1.0'")
    xmlDOM.appendChild node
'
	dim xmlRoot 
	set xmlRoot = xmlDOM.createElement( "EADATA" )
	xmlDOM.appendChild xmlRoot

	dim xmlDataSet
	set xmlDataSet = xmlDOM.createElement( "Dataset_0" )
	xmlRoot.appendChild xmlDataSet
	 
	dim xmlData 
	set xmlData = xmlDOM.createElement( "Data" )
	xmlDataSet.appendChild xmlData
	
	'add the contents for this scenario
	addScenarioContents xmlDom, xmlData, scenario

	getScenarioStepsXMLString = xmlDOM.xml
end function

function addScenarioContents(xmlDom, xmlData, scenario)
	'get the scenario xml
	dim scenarioXML
	set scenarioXML = CreateObject("MSXML2.DOMDocument")
	if not scenarioXML.LoadXML(scenario.XMLContent) then
		'exit if not a valid XML
		exit function
	end if
	'loop steps
	dim stepNodes
	set stepNodes = scenarioXML.SelectNodes("//step")
	dim stepNode
	for each stepNode in stepNodes
		'add details for each step
		addRow xmlDOM, xmlData, stepNode, scenarioXML
	next
end function

function addRow(xmlDOM, xmlData, stepNode, scenarioXML)
	
	dim xmlRow
	set xmlRow = xmlDOM.createElement( "Row" )
	xmlData.appendChild xmlRow
	
	'Step number
	dim xmlStep
	set xmlStep = xmlDOM.createElement( "step" )	
	xmlStep.text = stepNode.GetAttribute("level")
	xmlRow.appendChild xmlStep
	
	'Step Name
	dim xmlName
	set xmlName = xmlDOM.createElement( "name" )	
	xmlName.text = stepNode.GetAttribute("name")
	xmlRow.appendChild xmlName
	
	'Uses Name
	dim xmlUses
	set xmlUses = xmlDOM.createElement( "uses" )	
	xmlUses.text = stepNode.GetAttribute("uses")
	xmlRow.appendChild xmlUses
	
	'Requirement constraints
	dim xmlConstraints
	set xmlConstraints = xmlDOM.createElement("constraints")	
	xmlConstraints.text = getConstraintsText(xmlUses.text, scenarioXML)
	xmlRow.appendChild xmlConstraints
	
	'Expected Results
	dim xmlResult
	set xmlResult = xmlDOM.createElement( "result" )	
	xmlResult.text = stepNode.GetAttribute("result")
	xmlRow.appendChild xmlResult

	'State
	dim xmlState
	set xmlState = xmlDOM.createElement( "state" )	
	xmlState.text = stepNode.GetAttribute("state")
	xmlRow.appendChild xmlState
	
	
end function

function getConstraintsText(requirementName, scenarioXML)
	dim text
	text = ""
	getConstraintsText = text 'initialize
	'get the requirement GUID
	dim reqNode
	set reqNode = scenarioXML.SelectSingleNode("//item[@oldname = '" & sanitizeXMLString(requirementName) & "']")
	if reqNode is nothing then 
		exit function
	end if
	'get the requirement object
	dim requirementGUID 
	requirementGUID = reqNode.GetAttribute("guid")
	'get the constraint details with a query
	dim sqlGetData
	sqlGetData = "select oc.ConstraintType + ': ' + oc.[Constraint] as ConstraintText    " & vbNewLine & _
				" from t_objectconstraint oc                                            " & vbNewLine & _
				" inner join t_object o on o.Object_ID = oc.Object_ID                   " & vbNewLine & _
				" where o.ea_guid = '" & requirementGUID & "'                           " & vbNewLine & _
				" order by oc.Weight                                                    "
	dim constraintStrings
	set constraintStrings = getVerticalArrayListFromQuery(sqlGetData)
	dim constraintString
	if constraintStrings.Count = 0 then
		exit function
	end if
	text = Join(constraintStrings(0).ToArray(), vbNewLine)
	'return
	getConstraintsText = text
end function
