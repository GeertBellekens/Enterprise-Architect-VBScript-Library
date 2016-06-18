'[path=\Framework\Utils]
'[group=Utils]
'Author: Geert Bellekens
'Date: 2015-12-07
!INC Utils.Include

Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2

Class BinaryFile
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
		
	' Contents property, should be a array of bytes.
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
	
	'save the file to disk
	sub Save
		'first make sure the directory exists
		me.Folder.Save
		'then create file	
		'Create Stream object
		Dim BinaryStream
		Set BinaryStream = CreateObject("ADODB.Stream")
		'Specify stream type – we want To save binary data.
		BinaryStream.Type = adTypeBinary
		'Open the stream And write binary data To the object
		BinaryStream.Open
		BinaryStream.Write Contents
		'Save binary data To disk
		BinaryStream.SaveToFile me.FullPath, adSaveCreateOverWrite
	end sub
	
	'delete the file
	sub Delete
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		if fso.FileExists(me.FullPath) then
			fso.DeleteFile me.FullPath
		end if
	end sub

end class