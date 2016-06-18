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
	  me.Folder.FullPath = left(value, startBackslash -1) 'get everything before the last "\"
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
	

	
end class