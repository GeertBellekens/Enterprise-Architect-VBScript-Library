'[path=\Framework\Wrappers\Scripting]
'[group=Wrappers]

option explicit


!INC Utils.Include

'Author: Geert Bellekens
'Date: 2015-12-07

Class Script 
	Private m_Name
	Private m_Code
	Private m_Group
	Private m_Id

	Private Sub Class_Initialize
	  m_Name = ""
	  m_Code = ""
	  m_Id = ""
	  set m_Group = Nothing
	End Sub

	' Name property.
	Public Property Get Name
	  Name = m_Name
	End Property
	Public Property Let Name(value)
	  m_Name = value
	End Property

	' Code property.
	Public Property Get Code
	  Code = m_Code
	End Property
	Public Property Let Code(value)
	  m_Code = value
	End Property
	
	' Id property.
	Public Property Get Id
	  Id = m_Id
	End Property
	Public Property Let Id(value)
	  m_Id = value
	End Property
	
	' Path property.
	Public Property Get Path
	  Path = getPathFromCode
	  if len(Path) < 1 then
		Path = "\" & me.Group.Name
	  end if
	End Property

	' Group property.
	Public Property Get Group
	  set Group = m_Group
	End Property
	Public Property Let Group(value)
	  set m_Group = value
	  'add the script to the group
	   m_Group.Scripts.Add me
	End Property

	
	' Gets all scripts stored in the model
	Public function getAllScripts()
		dim queryResult
		dim resultArray
		dim row
		dim allGroups
		set allGroups = CreateObject("Scripting.Dictionary")
		dim allScripts
		set allScripts = CreateObject("System.Collections.ArrayList")
		dim sqlGet
		sqlGet = "select s.ScriptID, s.Notes, s.Script,ps.Script as SCRIPTGROUP, ps.Notes as GROUPNOTES, ps.ScriptID as GroupID " & _
					 " from t_script s " & _
					 " inner join t_script ps on s.ScriptAuthor = ps.ScriptName " & _
					 " where s.notes like '<Script Name=" & getWC() & "'"
        queryResult = Repository.SQLQuery(sqlGet)
		resultArray = convertQueryResultToArray(queryResult)
		dim id, notes, code, group, name, groupNotes, groupID, scriptGroup
		dim i
		For i = LBound(resultArray) To UBound(resultArray)
			notes = resultArray(i,2)
			id = resultArray(i,0)
			notes = resultArray(i,1)
			code = resultArray(i,2) 
			group = resultArray(i,3)
			groupNotes = resultArray(i,4)
			groupID = resultArray(i,5)
			if len(notes) > 0 then
				'first get or create the group
				if allGroups.Exists(groupID) then
					set scriptGroup = allGroups(groupID)
				else
					set scriptGroup = new ScriptGroup
					scriptGroup.Name = group
					scriptGroup.Id = groupID
					scriptGroup.setGroupTypeFromNotes groupNotes
					'add the group to the dictionary
					allGroups.Add groupID, scriptGroup
				end if
				'then make the script
				name = getNameFromNotes(notes)
				dim script
				set script = New Script
				script.Id = id
				script.Name = name
				script.Code = code
				'add the group to the script
				script.Group = scriptGroup
				'add the script to the list
				allScripts.Add script
			end if
		next
		set getAllScripts = allScripts
	End function
	
	'the notes contain= <Script Name="MyScriptName" Type="Internal" Language="JavaScript"/>
	'so the name is the second part when splitted by double quotes
	private function getNameFromNotes(notes)
		dim parts
		parts = split(notes,"""")
		getNameFromNotes = parts(1)
	end function
	
	'the path is defined in the code as '[path=\directory\subdirectory]
	private function getPathFromCode()
		dim returnPath
		returnPath = "" 'initialise emtpy
		dim pathIndicator, startPath, indPath
		pathIndicator = "[path="
		startPath = instr(me.Code, pathIndicator) + len(PathIndicator)
		if startPath > len(PathIndicator) then
			endPath = instr(startPath, me.Code, "]")
			if endPath > startPath then
				returnPath = mid(me.code,startPath, endPath - StartPath)
			end if
		end if
		getPathFromCode = returnPath
	end function
end Class