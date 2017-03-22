'[path=\Framework\Utils]
'[group=Utils]

'Name: CSVFile
'Author: Geert Bellekens
'Purpose: Wrapper script class for CSV files
'Date: 2017-03-20

!INC Utils.Include

'TODO: finish when actually needed...

Class CSVFile
	'private variables
	Private m_TextFile

	Private Sub Class_Initialize
		set m_TextFile = new TextFile
	End Sub
	
	
	' FileName property.
	Public Property Get FileName
	  FileName = m_TextFile.FileName
	End Property
	Public Property Let FileName(value)
	  m_TextFile.FileName = value
	End Property
	
	' Contents property. An ArrayList ArrayLists of strings
	Public Property Get Contents
	  Contents = m_TextFile.Contents
	End Property
	Public Property Let Contents(value)
	  m_TextFile.Contents = value
	End Property
	
end Class