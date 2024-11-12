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
	  set Contents = getDataFromCSVString(m_TextFile.Contents)
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
		m_TextFile.LoadContents
	End Property
	
	public Function getUserSelectedFileName()
		dim selectedFileName
		dim project
		set project = Repository.GetProjectInterface()
		m_TextFile.FullPath = project.GetFileNameDialog ("", "CSV files|*.csv", 1, 2 ,"", 1) 'save as with overwrite prompt: OFN_OVERWRITEPROMPT
	end function
	'save
	sub Save
		if len(me.FileName) = 0 then
			getUserSelectedFileName()
		end if		
		if len(me.FileName) = 0 then
			exit sub 'if still no path, then stop trying to save
		end if
		m_TextFile.Save
		Dim objStream
		Set objStream = CreateObject("ADODB.Stream")
		objStream.CharSet = "utf-8"
		objStream.Open
		objStream.WriteText m_TextFile.Contents
		objStream.SaveToFile me.FullPath, adSaveCreateOverWrite
	end sub
	
	private function getDataFromCSVString(csvstring)
		dim data
		set data = CreateObject("System.Collections.ArrayList")
		dim lines
		lines = split(csvstring, vbNewLine)
		dim line
		for each line in lines
			dim row
			set row = CreateObject("System.Collections.ArrayList")
			data.Add row
			dim parts
			parts = split(line, ",")
			dim part
			for each part in parts
				part = replace(part, """", "") 'remove all double quotes from the text
				row.Add part
			next
		next
		set getDataFromCSVString = data
	end function
	
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
			rowString = Join(row.ToArray(), ",")
			csvString = csvString & rowString
		next
		'return
		getCSVString = csvString
	end function
	
end Class