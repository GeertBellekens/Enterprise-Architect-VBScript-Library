'[path=\Framework\Wrappers\]
'[group=Wrappers]

!INC Utils.Include
!INC Local Scripts.EAConstants-VBScript

'
' Script Name: BusinessProcess
' Author: Geert Bellekens
' Purpose: Represents a BusinessProcess
' Date: 2022-01-05
'

dim g_AllBusinessProcesses 'global dictionary of all business processes using the GUID as key
set g_AllBusinessProcesses = CreateObject("Scripting.Dictionary")

'Business Processes should always be created witht this public createBusinessProcess statement to avoid creating bsuiness processes twice
Function getBusinessProcess(EAElement)
	if not g_AllBusinessProcesses.Exists(EAElement.ElementGUID) then
		dim newBusinessProcess
		set newBusinessProcess = New BusinessProcess
		newBusinessProcess.Init EAElement
	end if
	'return the element from the global dictionary
	set getBusinessProcess = g_AllBusinessProcesses(EAElement.ElementGUID)
end function

Class BusinessProcess
	Private m_calledProcesses
	Private m_wrappedElement
	Private m_allCalledProcesses
	Private m_allExchangedFisses
	Private m_allExchangedTechnicalMessages
	Private m_fisMessageCombos
	
	Private Sub Class_Initialize
	  set m_wrappedElement = nothing
	  set m_calledProcesses = nothing
	  set m_allCalledProcesses = nothing
	  set m_allExchangedFisses = nothing
	  set m_allExchangedTechnicalMessages = nothing
	  set m_fisMessageCombos = nothing
	End Sub
	
	Public sub Init(EAElement)
		if EAElement.Stereotype = "BusinessProcess" _
		  or EAElement.Stereotype = "Activity" then
			set m_wrappedElement = EAElement
			'add myself to the global list of business processes
			if not g_AllBusinessProcesses.Exists(me.GUID) then
				g_AllBusinessProcesses.Add me.GUID, me
			end if
		else
			Err.Raise vbObjectError + 1, "InvalidWrappedElement", "A BusinessProcess accepts only elements with steroetype 'BusinessProcess' or 'Activity'" 
		end if
	end sub
	
	public property Get wrappedElement
		set wrappedElement = m_wrappedElement
	end property
	
	public property Get GUID
		GUID = me.wrappedElement.ElementGUID
	end property
	
	public property Get name
		name = me.wrappedElement.Name
	end property
	
	
	public property Get calledProcesses
		if m_calledProcesses is nothing then
			getCalledProcesses
		end if
		set calledProcesses = m_calledProcesses
	end property
	
	public property Get allCalledProcesses
		if m_allCalledProcesses is nothing then
			getAllCalledProcesses
		end if
		set allCalledProcesses = m_allCalledProcesses
	end property
	
	'dictionary of all exchanged technical messages with the GUID as key
	public property Get allExchangedTechnicalMessages
		if m_allExchangedTechnicalMessages is nothing then
			retrieveFissesAndTechnicalMessages
		end if
		set allExchangedTechnicalMessages = m_allExchangedTechnicalMessages
	end property
	
	'dictionary of all exchanged fisses with the GUID as key
	public property Get allExchangedFisses
		if m_allExchangedFisses is nothing then
			retrieveFissesAndTechnicalMessages
		end if
		set allExchangedFisses = m_allExchangedFisses
	end property
	
	'dictionary of all combo's of fisses with their technical messages with the FISguid as key
	public property Get fisMessageCombos
		if m_fisMessageCombos is nothing then
			retrieveFissesAndTechnicalMessages
		end if
		set fisMessageCombos = m_fisMessageCombos
	end property
	
	'returns the fisses that correspond to this technical message
	public function getFissesForTechnicalMessage(technicalMessage)
		dim fisses
		set fisses = CreateObject("System.Collections.ArrayList")
		dim fisGUID
		for each fisGUID in me.fisMessageCombos.Keys
			if technicalMessage.ElementGUID = me.fisMessageCombos.Item(fisGUID).ElementGUID then
				fisses.add me.allExchangedFisses.Item(fisGUID)
			end if
		next
		'return
		set getFissesForTechnicalMessage = fisses
	end function
	
	private function retrieveFissesAndTechnicalMessages
		'initialize dictionaries
		set m_allExchangedFisses = CreateObject("Scripting.Dictionary")
		set m_allExchangedTechnicalMessages = CreateObject("Scripting.Dictionary")
		set m_FisMessageCombos = CreateObject("Scripting.Dictionary")
		'make a list of all GUID's of this process and of all the called processes
		dim guidList
		set guidList = CreateObject("System.Collections.ArrayList")
		'add own key
		guidList.Add me.GUID
		'add all called keys
		dim calledKey
		for each calledKey in me.allCalledProcesses.Keys
			guidList.Add calledKey
		next
		dim guidListString
		guidListString = Join(guidList.ToArray, "','")
		dim sqlGetData
		sqlGetData = "select distinct  fis.ea_guid as fisGUID, isnull(tmsg.ea_guid,'') as msgGUID                    " & vbNewLine & _
					" from t_connector c                                                                            " & vbNewLine & _
					" inner join t_connectortag tv on tv.ElementID = c.Connector_ID                                 " & vbNewLine & _
					" 							and tv.Property = 'MessageRef'                                      " & vbNewLine & _
					" inner join t_object fis on fis.ea_guid = tv.VALUE                                             " & vbNewLine & _
					" 					and fis.Stereotype = 'Message'                                              " & vbNewLine & _
					" 					and isnull(fis.Status, '') <> 'Not Needed'                                  " & vbNewLine & _
					" inner join t_object pl on pl.Object_ID in (c.Start_Object_ID, c.End_Object_ID)                " & vbNewLine & _
					" 					and pl.Stereotype = 'Pool'                                                  " & vbNewLine & _
					" inner join t_object bp on bp.Object_ID = pl.ParentID                                          " & vbNewLine & _
					" 					and bp.Stereotype in ('Activity', 'BusinessProcess')                        " & vbNewLine & _
					" 					and bp.ea_guid in ('" & guidListString & "')                                " & vbNewLine & _
					" left join (                                                                                   " & vbNewLine & _
					" 	select msg_fis.End_Object_ID, tmsg.ea_guid                                                  " & vbNewLine & _
					" 	from t_connector msg_fis                                                                    " & vbNewLine & _
					" 	inner join t_object msg on msg.Object_ID = msg_fis.Start_Object_ID                          " & vbNewLine & _
					" 								and msg.Stereotype = 'Message'                                  " & vbNewLine & _
					" 								and isnull(msg.Status, '') <> 'Not Needed'                      " & vbNewLine & _
					" 	inner join t_connector tmsg_msg on tmsg_msg.End_Object_ID = msg.Object_ID                   " & vbNewLine & _
					" 								and tmsg_msg.Connector_Type in ('Realization', 'Realisation')   " & vbNewLine & _
					" 	inner join t_object tmsg on tmsg.Object_ID = tmsg_msg.Start_Object_ID                       " & vbNewLine & _
					" 							and tmsg.Stereotype in ('XSDTopLevelElement', 'MA')                 " & vbNewLine & _
					" 							and isnull(tmsg.Status, '') <> 'Not Needed'                         " & vbNewLine & _
					" 	where msg_fis.Connector_Type in ('Realization', 'Realisation')                              " & vbNewLine & _
					" 	) tmsg on tmsg.End_Object_ID = fis.Object_ID                                                " & vbNewLine & _
					" where c.stereotype = 'MessageFlow'                                                            "
		dim results
		set results = getArrayListFromQuery(sqlGetData)
		dim row
		for each row in results
			'get FIS
			dim fisGUID
			fisGUID = row(0)
			dim fis as EA.Element
			if not m_allExchangedFisses.Exists(fisGUID) then
				set fis = Repository.GetElementByGuid(fisGUID)
				'put FIS in dictionary
				m_allExchangedFisses.Add fisGUID, fis
				'get message details
				dim msgGUID
				msgGUID = row(1)
				if len(msgGUID) > 0 then
					dim technicalMessage
					'get message object, first check if we already have it
					if not allExchangedTechnicalMessages.Exists(msgGUID) then
						set technicalMessage = Repository.GetElementByGuid(msgGUID)
						'add to dictionary
						m_allExchangedTechnicalMessages.add msgGUID, technicalMessage
					end if
					'get technical message
					set technicalMessage = m_allExchangedTechnicalMessages(msgGUID)
					'create combo
					m_FisMessageCombos.add fisGUID, technicalMessage
				end if
			else
				'this shoudn't happen becaue it means there are more than one technical messages for this fis
				'issue an error
				Repository.WriteOutput outPutName, now() & "ERROR: found multiple technical messages for FIS with GUID: '" & fisGUID  & "'", 0
			end if
		next
	end function
	
	private function getCalledProcesses
		set m_calledProcesses = CreateObject("Scripting.Dictionary")
		dim sqlGetData
		sqlGetData = "select distinct bp.Object_ID from t_object o                                  " & vbNewLine & _
					" inner join t_diagram d on d.ParentID = o.Object_ID                           " & vbNewLine & _
					" inner join t_diagramobjects do on do.Diagram_ID = d.Diagram_ID               " & vbNewLine & _
					" inner join t_object act on act.Object_ID = do.Object_ID                      " & vbNewLine & _
					" 						and act.Stereotype = 'Activity'                        " & vbNewLine & _
					" inner join t_objectproperties tv on tv.Object_ID = act.Object_ID             " & vbNewLine & _
					" 						and tv.Property = 'calledActivityRef'                  " & vbNewLine & _
					" inner join t_object bp on bp.ea_guid = tv.Value                              " & vbNewLine & _
					" 						and bp.Stereotype in ('Businessprocess', 'Activity')   " & vbNewLine & _
					" 						and isnull(bp.Status, '') <> 'Not Needed'              " & vbNewLine & _
					" inner join t_diagram bpd on bpd.ParentID = bp.Object_ID                      " & vbNewLine & _
					" where o.ea_guid = '"& me.GUID &"'                                            "
		dim elements
		set elements = getElementsFromQuery(sqlGetData)
		dim element as EA.Element
		for each element in elements
			dim subProcess
			set subProcess = getBusinessProcess(element)
			'add subProcess to list of calledProcesses
			m_calledProcesses.Add subProcess.GUID, subProcess
		next
	end function
	
	private function getAllCalledProcesses
		set m_allCalledProcesses = CreateObject("Scripting.Dictionary")
		'add direct called processes
		dim calledProcess
		for each calledProcess in me.calledProcesses.Items
			if not m_allCalledProcesses.Exists(calledProcess.GUID) then
				m_allCalledProcesses.Add calledProcess.GUID, calledProcess
				'add all called processes for this called process as well
				dim subCalledProcess
				for each subCalledProcess in calledProcess.allCalledProcesses.Items
					if not m_allCalledProcesses.Exists(subCalledProcess.GUID) then
						m_allCalledProcesses.Add subCalledProcess.GUID, subCalledProcess
					end if
				next
			end if
		next
	end function
	
end Class