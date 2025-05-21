'[path=\Framework\Utils]
'[group=Utils]

'Name: CSVFile
'Author: Geert Bellekens
'Purpose: Wrapper script class for CSV files
'Date: 2017-03-20

!INC Utils.Include


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
		'create CSV string
		m_TextFile.Contents = getCSVString(value)
	End Property
	
	' FullPath property.
	Public Property Get FullPath
	  FullPath = m_TextFile.FullPath
	End Property	
	public Property Let FullPath(value)
		m_TextFile.FullPath = value
	End Property
	
	'save
	sub Save
		m_TextFile.Save
'		Dim objStream
'		Set objStream = CreateObject("ADODB.Stream")
'		objStream.CharSet = "utf-8"
'		objStream.Open
'		objStream.WriteText me.Contents
'		objStream.SaveToFile me.FullPath, adSaveCreateOverWrite
	end sub
	
	private function getCSVString(arrayList)
		dim csvString
		csvString = ""
		dim row 'also an arrayList
		for each row in arrayList
			if len(csvString) > 0 then
				csvString = csvString & vbNewLine
			end if
			dim rowString
			'add double quotes for any value that isn't numerical
			dim value
			dim i
			for i = 0 to row.Count -1
				value = row(i)
				if not isNumeric(value) then
					row(i) = """" & value & """"
				end if
			next
			rowString = Join(row.ToArray(), ";")
			csvString = csvString & rowString
		next
		'return
		getCSVString = csvString
	end function
	
end Class