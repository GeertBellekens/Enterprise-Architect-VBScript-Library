'[path=\Framework\Wrappers\Scripting]
'[group=Wrappers]


!INC Utils.Include

'Author: Geert Bellekens
'Date: 2015-12-07
'constants for group type in database
Const gtNormal = "NORMAL", gtProjectBrowser = "PROJBROWSER", gtDiagram = "DIAGRAM", gtWorkflow = "WORKFLOW", _
  gtSearch = "SEARCH", gtModelSearch = "MODELSEARCH", gtContextElement = "CONTEXTELEMENT", _
  gtContextPackage = "CONTEXTPACKAGE", gtContextDiagram = "CONTEXTDIAGRAM", gtContextLink = "CONTEXTLINK"

'for some reason all groups have this value in column scriptCategory
Const scriptGroupCategory = "3955A83E-9E54-4810-8053-FACC68CD4782"

Class ScriptGroup 
	Private m_Id
	Private m_GUID
	Private m_Name
	Private m_GroupType
	Private m_Scripts
	
	Private Sub Class_Initialize
	  m_Id = ""
	  m_Name = ""
	  m_GroupType = ""
	  set m_Scripts = CreateObject("System.Collections.ArrayList")
	End Sub
	
	' Id property.
	Public Property Get Id
	  Id = m_Id
	End Property
	Public Property Let Id(value)
	  m_Id = value
	End Property	
	
	' GUID property.
	Public Property Get GUID
	  GUID = m_GUID
	End Property
	Public Property Let GUID(value)
	  m_GUID = value
	End Property	
	
	' Name property.
	Public Property Get Name
	  Name = m_Name
	End Property
	Public Property Let Name(value)
	  m_Name = value
	End Property

	' GroupType property.
	Public Property Get GroupType
	  GroupType = m_GroupType
	End Property
	Public Property Let GroupType(value)
	  m_GroupType = value
	End Property
	
	' Scripts property.
	Public Property Get Scripts
	  set Scripts = m_Scripts
	End Property
	
	'the notes contain something like <Group Type="NORMAL" Notes=""/>
	'so the group type is the second part when splitted by double quotes
	private function getGroupTypeFromNotes(notes)
		dim parts
		parts = split(notes,"""")
		getGroupTypeFromNotes = parts(1)
		if getGroupTypeFromNotes = "" then
			getGroupTypeFromNotes = gtNormal
		end if
	end function
	
	'sets the GroupType based on the given notes
	public sub setGroupTypeFromNotes(notes)
		GroupType = getGroupTypeFromNotes(notes)
	end sub
	
	'gets a dictionary of all groups without the scripts
	public function getAllGroups()
		dim allGroups, sqlGet
		dim queryResult
		dim resultArray
		set allGroups = CreateObject("Scripting.Dictionary")
		sqlGet = "select s.[ScriptID], s.[ScriptName] AS GroupGUID, s.[Notes], s.[Script] as GroupName " & _
				" from t_script s " & _
				" where s.Notes like '<Group Type=" & getWC() & "'"
		queryResult = Repository.SQLQuery(sqlGet)
		resultArray = convertQueryResultToArray(queryResult)
		dim groupId, groupGUID, groupName, notes, scriptGroup
		dim i
		For i = LBound(resultArray) To UBound(resultArray)
			groupId = resultArray(i,0)
			groupGUID = resultArray(i,1)
			notes = resultArray(i,2) 
			groupName = resultArray(i,3)
			if len(notes) > 0 then
				'first get or create the group
				if not allGroups.Exists(groupID) then
					set scriptGroup = new ScriptGroup
					scriptGroup.Name = groupName
					scriptGroup.Id = groupId
					scriptGroup.GUID = groupGUID
					scriptGroup.setGroupTypeFromNotes notes
					'add the group to the dictionary
					allGroups.Add groupID, scriptGroup
				end if
			end if
		next
		set getAllGroups = allGroups
	end function
	
	'Insert the group in the database
	public sub Create
		dim sqlInsert
		sqlInsert = "insert into t_script (ScriptCategory, ScriptName,Notes, Script) " & _
					" Values ('" & scriptGroupCategory & "','" & me.GUID & "','<Group Type=""" & me.GroupType & """ Notes=""""/>','" & me.Name & "')"
		Repository.Execute sqlInsert
	end sub

	public sub Update
		dim sqlUpdate
		sqlUpdate = "update t_script set Notes = '<Group Type=""" & me.GroupType & """ Notes=""""/>' where ScriptName = '" & me.GUID & "'"
		Repository.Execute sqlUpdate
	end sub

end Class