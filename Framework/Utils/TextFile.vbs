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