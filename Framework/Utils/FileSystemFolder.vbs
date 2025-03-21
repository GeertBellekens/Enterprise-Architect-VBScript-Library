'[path=\Framework\Utils]
'[group=Utils]

'Author: Geert Bellekens
'Date: 2015-12-07



!INC Local Scripts.EAConstants-VBScript
!INC Utils.Include

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
		Dim fso, result, folders, tempfolder, subfolder
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