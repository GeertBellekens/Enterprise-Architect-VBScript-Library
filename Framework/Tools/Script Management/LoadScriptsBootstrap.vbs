option explicit
'[path=\Framework\Tools\Script Management]
'[group=Script Management]

!INC Local Scripts.EAConstants-VBScript
!INC EAScriptLib.VBScript-GUID

' Author: Barrie Treloar
' Purpose: An all-in-one script, copied from the originals, to bootstrap the loads processes.
' 		   This should be a once-off-script, as it is not kept up-to-date.
'          Afterward use the LoadScripts script.
' Date: 2017-02-14

function getWC()
	if Repository.RepositoryType = "JET" then
		getWC = "*"
	else
		getWC = "%"
	end if
end function

function escapeSQLString(inputString)
	'replace the single quotes with two single quotes for all db types
	escapeSQLString = replace(inputString, "'","''")
	'dbspecifics
	select case Repository.RepositoryType
		case "POSTGRES"
			' replace backslash "\" by double backslash "\\"
			escapeSQLString = replace(escapeSQLString,"\","\\")
		case "JET"
			'replace pipe character | by '& chr(124) &'
			escapeSQLString = replace(escapeSQLString,"|", "'& chr(124) &'")
	end select
end function

Public Function convertQueryResultToArray(xmlQueryResult)
    Dim arrayCreated
    Dim i 
    i = 0
    Dim j 
    j = 0
    Dim result()
    Dim xDoc 
    Set xDoc = CreateObject( "MSXML2.DOMDocument" )
    'load the resultset in the xml document
    If xDoc.LoadXML(xmlQueryResult) Then        
		'select the rows
		Dim rowList
		Set rowList = xDoc.SelectNodes("//Row")

		Dim rowNode 
		Dim fieldNode
		arrayCreated = False
		'loop rows and find fields
		For Each rowNode In rowList
			j = 0
			If (rowNode.HasChildNodes) Then
				'redim array (only once)
				If Not arrayCreated Then
					ReDim result(rowList.Length, rowNode.ChildNodes.Length)
					arrayCreated = True
				End If
				For Each fieldNode In rowNode.ChildNodes
					'write f
					result(i, j) = fieldNode.Text
					j = j + 1
				Next
			End If
			i = i + 1
		Next
	end if
    convertQueryResultToArray = result
End Function

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

Class FileSystemFolder
	Private m_ParentPath
	Private m_Name
	
	' Name property.
	Public Property Get Name
	  Name = m_Name
	End Property
	Public Property Let Name(value)
	  m_Name = value
	End Property
	
	' FullPath property.
	Public Property Get FullPath
	  FullPath = m_ParentPath & "\" & Name
	End Property
	Public Property Let FullPath(value)
	  dim nameStart
	  nameStart = InstrRev(value, "\", -1, 0) 
	  m_ParentPath = left(value,nameStart -1)
	  m_Name = mid(value,NameStart +1)
	End Property
	
	'parentFolder
	Public Property Get ParentFolder
		set ParentFolder = nothing
		if len(m_ParentPath) > 0 and right(m_ParentPath,2) <> ":\" then
			set ParentFolder = new FileSystemFolder
			ParentFolder.FullPath = m_ParentPath
		end if
	End Property
	
	' TextFiles property
	Public Property Get TextFiles
		dim fso, fsoFolder, files, file, result, v_textFile, ts
		set result = CreateObject("System.Collections.ArrayList")
		Set fso = CreateObject("Scripting.FileSystemObject")
		if fso.FolderExists(me.FullPath) then
			Set fsoFolder = fso.GetFolder(me.FullPath)			
			Set files = fsoFolder.Files
			For Each file in files
				set v_textFile = new TextFile
				v_textFile.Folder = me
				v_textFile.FileName = file.Name
				set ts = file.OpenAsTextStream(ForReading, TristateUseDefault)
				v_textFile.Contents = ts.ReadAll
				ts.Close
				result.add v_textFile
			Next
		end if
		set TextFiles = result
	End Property
	
	'SubFolders property
	public property Get SubFolders
		Dim fso, result, folder, folders, tempfolder, subfolder
		set result = CreateObject("System.Collections.ArrayList")
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set tempfolder = fso.GetFolder(me.FullPath)
		Set folders = tempfolder.SubFolders
		For Each folder in folders
			set subfolder = new FileSystemFolder
			subFolder.FullPath = folder.Path
			result.Add subFolder
		Next
		set SubFolders = result
	End Property
	'let the user select a folder, optionally from a given starting path.
	public function getUserSelectedFolder(startPath)
		dim folder, shell
		Set shell  = CreateObject( "Shell.Application" )
		if len(startPath) > 0 then
			Set folder = shell.BrowseForFolder( 0, "Select Folder", 0,startPath)
		else
			Set folder = shell.BrowseForFolder( 0, "Select Folder", 0)
		end if
		if not folder is nothing then
			set getUserSelectedFolder = New FileSystemFolder
			getUserSelectedFolder.FullPath = folder.Self.Path 
			Session.Output "folder.Self.Path: " & folder.Self.Path
		else
			set getUserSelectedFolder = Nothing
		end if
	end function
	'save the folder
	public sub Save()
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		'first check if the path doesn't exist yet
		if not fso.FolderExists(me.FullPath) and len(me.FullPath) > 1 then
			if not me.ParentFolder is nothing then
				me.ParentFolder.Save
			end if
			fso.CreateFolder me.FullPath
		end if
	end sub
	'delete the folder
	public sub Delete()
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		if fso.FolderExists(me.FullPath) then
			fso.DeleteFolder me.FullPath
		end if
	end sub
End Class

Const scriptCategory = "605A62F7-BCD0-4845-A8D0-7DC45B4D2E3F"

Class TextFile
	Private m_FullPath
	Private m_Contents
	Private m_Folder
	Private m_FileName

	Private Sub Class_Initialize
	  set m_Folder = Nothing
	  m_FileName = ""
	  m_Contents = ""
	End Sub
	
	' FullPath property.
	Public Property Get FullPath
	  FullPath = me.Folder.FullPath & "\" & me.FileName
	End Property	
	public Property Let FullPath(value)
	  dim startBackslash
	  startBackslash = InstrRev(value, "\", -1, 1)
	  dim folderPath
	  folderPath = left(value, startBackslash -1) 'get everything before the last "\"
	  if ucase(folderPath) <> ucase(me.Folder.FullPath) then
		'make new folder object to avoid side effects on the folder object
		me.Folder = New FileSystemFolder
		me.Folder.FullPath = left(value, startBackslash -1) 'get everything before the last "\"
	  end if
	  me.FileName = mid(value, startBackslash + 1) 'get everything after the last "."
	end Property
		
	' Contents property.
	Public Property Get Contents
	  Contents = m_Contents
	End Property
	Public Property Let Contents(value)
	  m_Contents = value
	End Property
	
	' FileName property.
	Public Property Get FileName
	  FileName = m_FileName
	End Property
	Public Property Let FileName(value)
	  m_FileName = value
	End Property
	' FileNameWithoutExtension property.
	Public Property Get FileNameWithoutExtension
	  dim startExtension
	  startExtension = InstrRev(me.FileName, ".", -1, 1)
	  FileNameWithoutExtension = left(me.FileName, startExtension -1) 'get everything before the last "."
	End Property
	' Extension property.
	Public Property Get Extension
	  dim startExtension
	  startExtension = InstrRev(me.FileName, ".", -1, 1)
	  Extension = mid(me.FileName, startExtension + 1) 'get everything after the last "."
	End Property
	
	' Folder property.
	Public Property Get Folder
	  if m_Folder is nothing then
		set m_Folder = new FileSystemFolder
	  end if
	  set Folder = m_Folder
	End Property
	Public Property Let Folder(value)
	  set m_Folder = value
	End Property
	
	'save the file
	sub Save
		Dim fso, MyFile
		Set fso = CreateObject("Scripting.FileSystemObject")
		'first make sure the directory exists
		me.Folder.Save
		'then create file
		Set MyFile = fso.CreateTextFile(me.FullPath, True)
		MyFile.Write(Contents)
		MyFile.close
	end sub
	
	'delete the file
	sub Delete
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		if fso.FileExists(me.FullPath) then
			fso.DeleteFile me.FullPath
		end if
	end sub
	'let the user select a file from the file system
	public function UserSelect(initialDir,filter)
		dim selectedFileName
		selectedFileName = ChooseFile(initialDir,filter)
		'check if anything was selected
		if len(selectedFileName) > 0 then
			me.FullPath = selectedFileName
			UserSelect = true
			me.LoadContents
		else
			UserSelect = false
		end if
	end function
	'load the contents of the file from the file system
	public function loadContents()
		Dim fso
		dim fsoFile
		dim ts
		Set fso = CreateObject("Scripting.FileSystemObject")
		if fso.FileExists(me.FullPath) then
			set fsoFile = fso.GetFile(me.FullPath)
			set ts = fsoFile.OpenAsTextStream(ForReading, TristateUseDefault)
			me.Contents = ts.ReadAll
		end if
	end function
	
end class
Class Script 
	Private m_Name
	Private m_Code
	Private m_Group
	Private m_Id
	Private m_GUID

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
	
	' GUID property.
	Public Property Get GUID
	  GUID = m_GUID
	End Property
	Public Property Let GUID(value)
	  m_GUID = value
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
	
	' GroupNameInCode property
	Public Property Get GroupInNameCode
	  GroupInNameCode = getGroupFromCode()
	End Property

	
	' Gets all scripts stored in the model
	Public function getAllScripts(allGroups)
		dim resultArray, scriptGroup,row,queryResult
		set scriptGroup = new scriptGroup
		set allGroups = scriptGroup.getAllGroups()
		dim allScripts
		set allScripts = CreateObject("System.Collections.ArrayList")
		dim sqlGet
		sqlGet = "select s.ScriptID, s.Notes, s.Script,ps.Script as SCRIPTGROUP, ps.Notes as GROUPNOTES, ps.ScriptID as GroupID, ps.ScriptName as GroupGUID, s.ScriptName as ScriptGUID " & _
					 " from t_script s " & _
					 " inner join t_script ps on s.ScriptAuthor = ps.ScriptName " & _
					 " where s.notes like '<Script Name=" & getWC() & "'"
        queryResult = Repository.SQLQuery(sqlGet)
		resultArray = convertQueryResultToArray(queryResult)
		dim id, notes, code, group, name, groupNotes, groupID, groupGUID, scriptGUID
		dim i
		For i = LBound(resultArray) To UBound(resultArray)
			id = resultArray(i,0)
			notes = resultArray(i,1)
			code = resultArray(i,2) 
			group = resultArray(i,3)
			groupNotes = resultArray(i,4)
			groupID = resultArray(i,5)
			groupGUID = resultArray(i,6)
			scriptGUID = resultArray(i,7)
			if len(notes) > 0 then
				'first get or create the group
				if allGroups.Exists(groupID) then
					set scriptGroup = allGroups(groupID)
				else
					set scriptGroup = new ScriptGroup
					scriptGroup.Name = group
					scriptGroup.Id = groupID
					scriptGroup.GUID = groupGUID
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
				script.GUID = scriptGUID
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
		getPathFromCode = getKeyValue("path")
	end function
	'the Group is defined in the code as '[group=NameOfTheGroup]
	public function getGroupFromCode()
		getGroupFromCode = getKeyValue("group")
	end function
	
	'the key-value pair is defined in the code as '[keyName=value]
	public function getKeyValue(keyName)
		dim returnValue
		returnValue = "" 'initialise emtpy
		dim keyIndicator, startKey, endKey, tempValue
		keyIndicator = "[" & keyName & "=" 
		startKey = instr(me.Code, keyIndicator) + len(keyIndicator)
		if startKey > len(keyIndicator) then
			endKey = instr(startKey, me.Code, "]")
			if endKey > startKey then
				tempValue = mid(me.code,startKey, endKey - startKey)
				'filter out newline in case someone forgot to add the closing "]"
				if instr(tempValue,vbNewLine) = 0 and instr(tempValue,vbLF) = 0 then
					returnValue = tempValue
				end if
			end if
		end if
		getKeyValue = returnValue
	end function
	
	public function addGroupToCode()
		dim groupFromCode
		groupFromCode = me.getGroupFromCode()
		if not len(groupFromCode) > 0 then
			'add the group indicator
			me.Code = "'[group=" & me.Group.Name & "]" & vbNewLine & me.Code
		end if
	end function
	
	
	'Insert the script in the database
	public sub Create
		dim sqlInsert
		sqlInsert = "insert into t_script (ScriptCategory, ScriptName, ScriptAuthor, Notes, Script) " & _
					" Values ('" & scriptCategory & "','" & me.GUID & "','" & me.Group.GUID & "','<Script Name=""" & me.Name & """ Type=""Internal"" Language=""VBScript""/>','" & escapeSQLString(me.Code) & "')"
'		Session.Output "***********************************************"
'		Session.Output "sql = " & sqlInsert			
		Repository.Execute sqlInsert
		If Err.Number <> 0 Then
			Session.Output "Error: " & Err.Number
			Session.Output "Error (Hex): " & Hex(Err.Number)
			Session.Output "Source: " &  Err.Source
			Session.Output "Description: " &  Err.Description
			Err.Clear
		End If
		
	end sub
	
	'update the script in the database
	public sub Update
		dim sqlUpdate
		sqlUpdate = "update t_script set script = '" & escapeSQLString(me.Code) & "', ScriptAuthor = '" & me.Group.GUID & _
					"', Notes = '<Script Name=""" & me.Name & """ Type=""Internal"" Language=""VBScript""/>' where ScriptName = '" & me.GUID & "'"
		Repository.Execute sqlUpdate
		If Err.Number <> 0 Then
			Session.Output "Error: " & Err.Number
			Session.Output "Error (Hex): " & Hex(Err.Number)
			Session.Output "Source: " &  Err.Source
			Session.Output "Description: " &  Err.Description
			Err.Clear
		End If

	end sub
	
end Class

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
		If Err.Number <> 0 Then
			Session.Output "Error: " & Err.Number
			Session.Output "Error (Hex): " & Hex(Err.Number)
			Session.Output "Source: " &  Err.Source
			Session.Output "Description: " &  Err.Description
			Err.Clear
		End If
		
	end sub

end Class

' Author: Geert Bellekens
' Purpose: Loads scripts from the file systems and stores them in Enterprise Architect
' Date: 2015-12-07
'

'gets all the scripts from the given folder and its subfolders (if any)
function getScriptsFromFolder(selectedFolder, allGroups, allScripts, overwriteExisting)
	dim script, subFolder, file
	for each file in selectedFolder.TextFiles
		Session.Output "FileName: " & file.FileName
		'Session.Output "Code: " & file.Contents
		set script = getScriptFromFile(file, allGroups, allScripts,overwriteExisting)
		if overwriteExisting = vbCancel then
			exit for
		end if
	next
	'then process subfolders
	if not overwriteExisting = vbCancel then
		for each subFolder in selectedFolder.SubFolders
			getScriptsFromFolder subFolder, allGroups, allScripts, overwriteExisting
		next
	end if
end function

function getScriptFromFile(file, allGroups, allScripts,overwriteExisting)
	dim script, newScript, foundMatch, newScriptGroupName, group, foundGroup
	foundMatch = false
	foundGroup = false
	set group = nothing
	set script = Nothing
	if file.Extension = "vbs" then
		for each script in allScripts
			set newScript = new Script
			newScript.Name = file.FileNameWithoutExtension
			newScript.Code = file.Contents
			newScriptGroupName = newScript.GroupInNameCode 
			'if the groupname was not found in the code we use the name of the package
			if len(newScriptGroupName) = 0 then
				newScriptGroupName = file.Folder.Name
			end if
			'check the name of the script
			if script.Name = newScript.Name then
				'check if there is a groupname defined in the file
				if script.Group.Name = newScriptGroupName then
					'we have a match
					foundMatch = true
					set group = script.Group
					exit for
				end if
			end if
		next
		if not foundMatch then 
			'script did not exist yet
			'figure out if the group exists already
			for each group in allGroups.Items
				if group.Name = newScriptGroupName then
					'found the group
					'add the group to the new script
					newScript.Group = group
					foundGroup = true
					exit for
				end if
			next
			'if the group doesnt exist yet we have to create it
			if not foundGroup then
				set group = new ScriptGroup
				group.Name = newScriptGroupName
				group.GUID = GUIDGenerateGUID()
				group.GroupType = gtNormal
				'create the Group in the database
				group.Create
				'refresh allGroups
				Session.Output "allGroups.Count before: " & allGroups.Count
				set allGroups = group.GetAllGroups()
				Session.Output "allGroups.Count after: " & allGroups.Count
				'add the group to the script
				newScript.Group = group
			end if
			'Now we have to create the script
			newScript.GUID = GUIDGenerateGUID()
			newScript.Create
			set script = newScript
		else
			if overwriteExisting = "undecided" then
				overwriteExisting = Msgbox("Do you want to update existing scripts?", vbYesNoCancel+vbQuestion, "Update existing scripts")
			end if
			if overwriteExisting = vbYes then
				script.Code = newScript.Code
				script.Update
			end if
		end if

	end if
	set getScriptFromFile = script
end function

sub main
	dim selectedFolder,file, allScripts, allGroups,script, overwriteExisting
	set selectedFolder = new FileSystemFolder
	set selectedFolder = selectedFolder.getUserSelectedFolder("C:\SparxEA-Scripts\Enterprise-Architect-VBScript-Library\Framework")
	overwriteExisting = "undecided"
	if not selectedFolder is nothing then
		set allGroups = Nothing
		set script = new Script
		'first get all existing scripts and groups
		set allScripts = Script.getAllScripts(allGroups)
		'get the scripts from the folder and its subfolders
		getScriptsFromFolder selectedFolder, allGroups, allScripts, overwriteExisting
	end if
end sub

sub main1
	dim a_script, folder, file, files, scriptGroup, allGroups, allScripts, overwriteExisting
	
	set scriptGroup = new scriptGroup
    set allGroups = scriptGroup.getAllGroups()
	set a_script = new Script
	set allScripts = a_script.getAllScripts(allGroups)
	overwriteExisting = vbYes

	set folder = New FileSystemFolder
	folder.FullPath = "C:\SparxEA-Scripts\Enterprise-Architect-VBScript-Library\Framework\Tools\Script Management"
	
	set files = folder.TextFiles
	for each file in files
'		Session.Output "------------------------------------------"
'		Session.Output "FileName: " & file.FileName
'		Session.Output "File Contents"
'		Session.Output file.Contents
'		Session.Output "=========================================="
'		Session.Output "Escaped Contents"
'		Session.Output escapeSQLString(file.Contents)
'		Session.Output "$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"
		set a_script = getScriptFromFile(file, allGroups, allScripts, overwriteExisting)
'		Session.Output "------------------------------------------"
	next
	
	If Err.Number <> 0 Then
	    Session.Output "Error: " & Err.Number
		Session.Output "Error (Hex): " & Hex(Err.Number)
		Session.Output "Source: " &  Err.Source
		Session.Output "Description: " &  Err.Description
		Err.Clear
	End If
end sub

main