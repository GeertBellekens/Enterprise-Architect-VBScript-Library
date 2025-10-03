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
	Private m_ScenarioSteps
	Private m_XmlDom
	private m_MainScenario

	Private Sub Class_Initialize
		m_Name			= ""
		m_Notes 		= ""
		m_ScenarioType 	= ""
		m_GUID        	= ""
		m_Entry       	= ""
		m_Join        	= ""
		m_XMLContent  	= ""
		set m_ScenarioSteps = CreateObject("System.Collections.ArrayList")
		set m_XmlDom = CreateObject("MSXML2.DOMDocument")
		set m_MainScenario = nothing
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

	Public Property Get EntryLevel
		EntryLevel = 0 'default
		If Len(Me.Entry) > 1 Then
			'get all but the last character of the Entry to get the level
			EntryLevel = CInt(Left(Me.Entry, Len(Me.Entry) - 1))
		End If
	End Property
	Public Property Get ExtensionLetter
		ExtensionLetter = ""
		if len(Me.Entry) > 1 then
			ExtensionLetter = right(me.Entry, 1)
		end if
	end Property

	' Join property.
	Public Property Get Join
		Join = m_Join
	End Property
	Public Property Let Join(value)
		m_Join = value
	End Property
	'JoinLevel
	Public Property Get JoinLevel
		JoinLevel = 0 'default
		if len(me.Join) > 0 then
			JoinLevel = Cint(me.Join)
		end if
	end Property

	' XMLContent property.
	Public Property Get XMLContent
		XMLContent = m_XMLContent
	End Property
	Public Property Let XMLContent(value)
		m_XMLContent = value
		Me.XmlDom.LoadXML value
	End Property

	' ScenarioSteps property
	Public Property Get ScenarioSteps
		Set ScenarioSteps = m_ScenarioSteps
	End Property
	
	' XmlDom property
	Public Property Get XmlDom
		set XmlDom = m_XmlDom
	End Property
	
	' MainScenario property
	Public Property Get MainScenario
		Set MainScenario = m_MainScenario
	End Property
	Public Property Let MainScenario(value)
		Set m_MainScenario = value
		me.ResolveEntryAndJoin
	End Property
	
	
	
	public function initialize(name, scenarioType, xmlContent, notes, guid)
		me.Name = name
		me.ScenarioType = scenarioType
		me.XMLContent = xmlContent
		me.Notes = notes
		me.GUID = guid
		'initialize the steps of the scenario
		initializeScenarioSteps
	end function
	
	function initializeScenarioSteps()
		dim stepNodes
		set stepNodes = me.XmlDom.selectNodes("//step")
		dim stepNode
		for each stepNode in stepNodes
			'create new scenario step
			dim scenarioStep
			set scenarioStep = new UsecaseScenarioStep
			'initialize the step
			scenarioStep.initialize stepNode
			'add to list
			me.ScenarioSteps.Add scenarioStep
		next
	end function
	
	public function ResolveEntryAndJoin()
		if me.MainScenario is nothing then
			exit function
		end if
		dim extensionNode
		set extensionNode = me.MainScenario.XmlDom.SelectSingleNode("//extension[@guid = '" & me.GUID & "']")
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
			set joinStepNode = me.MainScenario.XmlDom.SelectSingleNode("//step[@guid = '" & joinStepGUID & "']")
			if not joinStepNode is nothing then
				me.Join = joinStepNode.GetAttribute("level")
			end if
		end if
	end function
	
	function getCombinedScenariosteps()
		dim altCounter
		altCounter = 0
		if me.ScenarioType = ucAlternate then
			altCounter = 1
		end if
		'return regular scenariosteps if main scenario is not defined
		if me.MainScenario is nothing then
			set getCombinedScenariosteps = me.ScenarioSteps
			exit function
		end if
		dim combinedScenarioSteps
		set combinedScenarioSteps = CreateObject("System.Collections.ArrayList")
		'get steps before from main scenario
		dim i
		dim mainScenarioSteps
		set mainScenarioSteps = me.MainScenario.ScenarioSteps
		for i = 0 to me.EntryLevel - 1 - altCounter 
			if mainScenarioSteps.Count > i then
				dim tempStep
				set tempStep = mainScenarioSteps(i)
				combinedScenarioSteps.Add tempStep
			end if
		next
		'add own steps with extension letter.
		dim extensionLetter
		extensionLetter = Me.ExtensionLetter
		dim scenarioStep
		for each scenarioStep in me.ScenarioSteps
			scenarioStep.Level = (scenarioStep.Level + me.EntryLevel - altCounter) & extensionLetter
			combinedScenarioSteps.Add scenarioStep
		next
		'get steps after
		if me.JoinLevel > 0 then
			for i = me.JoinLevel - 1 to mainScenarioSteps.Count -1
				combinedScenarioSteps.Add mainScenarioSteps(i)
			next
		end if
		'return
		set getCombinedScenariosteps = combinedScenarioSteps
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
	dim mainScenario
	set mainScenario = nothing
	for each scenarioRow in results
		dim scenario
		set scenario = new UsecaseScenario
		scenario.initialize scenarioRow(0), scenarioRow(1), scenarioRow(2), scenarioRow(3), scenarioRow(4)
		scenarios.Add scenario
		'the main scenario xml contains the entry and join information
		if mainScenario is nothing then
			set mainScenario = scenario
		else
			scenario.MainScenario = mainScenario
		end if
	next
	'return
	set getScenariosForUseCase = scenarios
end function

