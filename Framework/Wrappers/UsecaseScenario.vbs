'[path=\Framework\Wrappers]
'[group=Wrappers]



!INC Utils.Include

'Author: Geert Bellekens
'Date: 2023-03-03

const ucBasicPath = "Basic Path"
const ucAlternate = "Alternate"
const ucException = "Exception"

Class UsecaseScenario 
	Private m_Name
	Private m_Notes
	Private m_ScenarioType
	Private m_GUID
	Private m_Entry
	Private m_Join
	Private m_XMLContent

	Private Sub Class_Initialize
		m_Name        = ""
		m_Notes = ""
		m_ScenarioType = ""
		m_GUID        = ""
		m_Entry       = ""
		m_Join        = ""
		m_XMLContent  = ""
	End Sub

	' Name property.
	Public Property Get Name
	  Name = m_Name
	End Property
	Public Property Let Name(value)
	  m_Name = value
	End Property
	' GUID property.
	Public Property Get GUID
	  GUID = m_GUID
	End Property
	Public Property Let GUID(value)
	  m_GUID = value
	End Property	
	' Notes property.
	Public Property Get Notes
	  Notes = m_Notes
	End Property
	Public Property Let Notes(value)
	  m_Notes = value
	End Property
	' ScenarioType property.
	Public Property Get ScenarioType
	  ScenarioType = m_ScenarioType
	End Property
	Public Property Let ScenarioType(value)
	  m_ScenarioType = value
	End Property
	' Entry property.
	Public Property Get Entry
	  Entry = m_Entry
	End Property
	Public Property Let Entry(value)
	  m_Entry = value
	End Property
	' Join property.
	Public Property Get Join
	  Join = m_Join
	End Property
	Public Property Let Join(value)
	  m_Join = value
	End Property
	' XMLContent property.
	Public Property Get XMLContent
	  XMLContent = m_XMLContent
	End Property
	Public Property Let XMLContent(value)
	  m_XMLContent = value
	End Property
	
	public function initialize(name, scenarioType, xmlContent, notes, guid)
		me.Name = name
		me.ScenarioType = scenarioType
		me.XMLContent = xmlContent
		me.Notes = notes
		me.GUID = guid
	end function
	
	public function ResolveEntryAndJoin(mainScenarioXMLDom)
		dim extensionNode
		set extensionNode = mainScenarioXMLDom.SelectSingleNode("//extension[@guid = '" & me.GUID & "']")
		if extensionNode is nothing then
			'no entry or join found in main scenario
			exit function
		end if
		me.Entry = extensionNode.GetAttribute("level")
		dim joinStepGUID
		joinStepGUID = extensionNode.GetAttribute("join")
		'fin the step with the given joinStepGUID
		if len(joinStepGUID) > 0 then
			dim joinStepNode
			set joinStepNode = mainScenarioXMLDom.SelectSingleNode("//step[@guid = '" & joinStepGUID & "']")
			if not joinStepNode is nothing then
				me.Join = joinStepNode.GetAttribute("level")
			end if
		end if
	end function
end Class

function getScenariosForUseCase(useCaseObjectID)
	dim scenarios
	set scenarios = CreateObject("System.Collections.ArrayList")
	dim sqlGetData
	sqlGetData = "select oc.Scenario, oc.ScenarioType, oc.XMLContent, oc.Notes, oc.ea_guid          " & vbNewLine & _
				" from t_objectscenarios oc                                                        " & vbNewLine & _
				" inner join t_object o on o.Object_ID = oc.Object_ID                              " & vbNewLine & _
				" where o.Object_ID = " & useCaseObjectID & "                                      " & vbNewLine & _
				" order by case when oc.ScenarioType = 'Basic Path' then 1 else 2 end, oc.EValue   "
	dim results
	set results = getArrayListFromQuery(sqlGetData)
	dim scenarioRow
	dim mainScenarioXMLDom
	set mainScenarioXMLDom = nothing
	for each scenarioRow in results
		dim scenario
		set scenario = new UsecaseScenario
		scenario.initialize scenarioRow(0), scenarioRow(1), scenarioRow(2), scenarioRow(3), scenarioRow(4)
		scenarios.Add scenario
		'the main scenario xml contains the entry and join information
		if mainScenarioXMLDom is nothing then
			set mainScenarioXMLDom = CreateObject("MSXML2.DOMDocument")
			mainScenarioXMLDom.LoadXML scenarioRow(2)
		else
			scenario.ResolveEntryAndJoin mainScenarioXMLDom
		end if
	next
	'return
	set getScenariosForUseCase = scenarios
end function

