'[path=\Framework\Utils]
'[group=Utils]
'Author: Geert Bellekens
'Date: 2015-12-07
!INC Utils.Include


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
          FileNameWithoutExtension = FileName
	  startExtension = InstrRev(me.FileName, ".", -1, 1)
          if startExtension <> 0 then
	    FileNameWithoutExtension = left(me.FileName, startExtension -1) 'get everything before the last "."
          end if
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
		'invalid characters cause errors, ignore them
		On Error Resume Next
		'first make sure the directory exists
		me.Folder.Save
		If Err.Number <> 0 Then
			Session.Output "ERROR: " &  Err.Description
			Err.Clear
			exit sub
		End If

		'then create file
		Set MyFile = fso.CreateTextFile(me.FullPath, True)
		If Err.Number <> 0 Then
			Session.Output "ERROR: " &  Err.Description
			Err.Clear
			exit sub
		End If
		On Error Goto 0

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
		Dim BinaryStream
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		if fso.FileExists(me.FullPath) then
			' Load with ADO stream to handle UTF-8
			' See https://github.com/GeertBellekens/Enterprise-Architect-VBScript-Library/issues/25
			Set BinaryStream = CreateObject("ADODB.Stream")
			With BinaryStream
				.Charset = "utf-8"
				.Type = 2
				.Open
				.LoadFromFile me.FullPath
				me.Contents = BinaryStream.ReadText()
				.Close
			End With
		end if
	end function
	'appends the given string to the end of the textfile
	public function append(contentToAppend)
		dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		dim fsoFile
		if not fso.FileExists(me.FullPath) then
			'create as new file
			me.Contents = contentToAppend
			me.save
		else
			'then append to the file
			Set fsoFile = fso.OpenTextFile(me.FullPath, ForAppending,TristateUseDefault)
			fsoFile.Write contentToAppend
			fsoFile.Close
		end if
	end function
	
end class

'Static functions
function writeFile(filename, contents)
	dim file
	set file = New TextFile
	file.FullPath = filename
	file.Contents = contents
	file.Save
end function