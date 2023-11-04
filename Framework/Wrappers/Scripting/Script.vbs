'[path=\Framework\Wrappers\Scripting]
'[group=Wrappers]



!INC Utils.Include

'Author: Geert Bellekens
'Date: 2015-12-07

'for some reason all scripts in the database have this value in the column scriptCategory
Const scriptCategory = "605A62F7-BCD0-4845-A8D0-7DC45B4D2E3F"

Class Script
        Private MAX_SQL_SIZE

        Private m_Name
        Private m_Code
        Private m_Group
        Private m_Id
        Private m_GUID

        Private Sub Class_Initialize
          ' These files fail
          '   74556 ./Projects/Project A/Reports and exports/FIS - BP - Matrix export.vbs
          '  293828 ./Framework/Tools/Script Management/LoadScriptsBootstrap.vbs
          MAX_SQL_SIZE = 75000
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

        ' GroupTypeFromCode property
        Public Property Get GroupType
                GroupType = getGroupTypeFromCode()
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
        ' the Group Type is defined in the code as '[group_type=GroupType]
        ' if not specified then defaults to gtNormal
        private function getGroupTypeFromCode()
                getGroupTypeFromCode = getKeyValue("group_type")
                if getGroupTypeFromCode = "" then
                        getGroupTypeFromCode = gtNormal
                end if
        end function

        'the key-value pair is defined in the code as '[keyName=value]
        public function getKeyValue(keyName)
                dim returnValue
                returnValue = "" 'initialise emtpy
                getKeyValue = returnValue

                dim keyIndicator, startKey, endKey, tempValue, prevCharIndex, prevChar
                keyIndicator = "'[" & keyName & "="
                'Session.Output "DEBUG: getKeyValue: searching for " & keyIndicator
                startKey = instr(me.Code, keyIndicator)
                if startKey = 0 then
                        'no matches found
                        'Session.Output "DEBUG: getKeyValue: no start key found, " & keyIndicator
                        'Session.Output "DEBUG: getKeyValue: in Script " & me.Name
                        'Session.Output "DEBUG: getKeyValue: code=" & me.Code
                        exit function
                end if

                'only allow key/values at the beginning of a line
                prevCharIndex = startKey - 1
                'check at beginning of line, prevCharIndex=0 indicates its the first line of the file
                if prevCharIndex <> 0 then
                        prevChar = mid(me.Code, prevCharindex, 1)
                        if prevChar <> vbLf and prevChar <> vbCr then
                                'key/value not at start of line
                                'Session.Output "DEBUG: getKeyValue: previous char is not new line, '" & Asc(prevChar) & "' at index " & prevCharindex
                                exit function
                        end if
                end if

                endKey = instr(startKey, me.Code, "]")
                if endKey = 0 then
                        'no closing bracket
                        'Session.Output "DEBUG: getKeyValue: no closing bracket ']'"
                        exit function
                end if

                'move startkey to after keyIndicator
                startKey = startKey + len(keyIndicator)
                tempValue = mid(me.code, startKey, endKey - startKey)
                'filter out multiline results in case someone forgot to add the closing "]"
                if instr(tempValue,vbNewLine) <> 0 or instr(tempValue,vbLF) <> 0 then
                        Session.Output "ERROR: " & keyIndicator & " missing closing ] on same line"
                        exit function
                end if
                returnValue = tempValue

                getKeyValue = returnValue
                'Session.Output "DEBUG: getKeyValue: " & keyName & " = " & getKeyValue
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
                Session.Output "DEBUG: Script.Create Len(sqlInsert)=" & Len(sqlInsert)
                if Len(sqlInsert) < MAX_SQL_SIZE then
                   Repository.Execute sqlInsert
                else
                   Session.Output "ERROR: Unable to create script " & me.Group.Name & "." & me.Name & " due to database limits"
                end if
        end sub

        'update the script in the database
        public sub Update
                dim sqlUpdate
                sqlUpdate = "update t_script set script = '" & escapeSQLString(me.Code) & "', ScriptAuthor = '" & me.Group.GUID & _
                                        "', Notes = '<Script Name=""" & me.Name & """ Type=""Internal"" Language=""VBScript""/>' where ScriptName = '" & me.GUID & "'"
                Session.Output "DEBUG: Script.Update Len(sqlUpdate)=" & Len(sqlUpdate)
                if Len(sqlUpdate) < MAX_SQL_SIZE then
                   Repository.Execute sqlUpdate
                else
                   Session.Output "ERROR: Unable to create script " & me.Group.Name & "." & me.Name & " due to database limits"
                end if
        end sub

end Class