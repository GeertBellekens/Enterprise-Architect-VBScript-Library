'[path=\Framework\Wrappers\Scripting]
'[group=Wrappers]

!INC Utils.Include

'Author: Geert Bellekens
'Date: 2015-12-07

Class SearchResults
'#region private attributes
	private m_Fields
	private m_Results
	private m_Name
'#endregion private attributes

'#region "Constructor"
	Private Sub Class_Initialize
		me.Fields = CreateObject("System.Collections.ArrayList")
		me.Results = CreateObject("System.Collections.ArrayList")
		me.Name = ""
	end sub
'#endregion "Constructor"
	
'#region Properties

	' Fields property
	Public Property Get Fields
	  set Fields = m_Fields
	End Property
	Public Property Let Fields(value)
	  set m_Fields = value
	End Property
	
	' Results property
	Public Property Get Results
	  set Results = m_Results
	End Property
	Public Property Let Results(value)
	  set m_Results = value
	End Property
	
	' Name property.
	Public Property Get Name
	  Name = m_Name
	End Property
	Public Property Let Name(value)
	  m_Name = value
	End Property	
'#endregion Properties
	
'#region functions
	'Show this resultset in the model search window
	public function Show()
		Repository.RunModelSearch me.Name,"searchTerm","searchOptions", makeSearchDataString()
	end function
	
	private function makeSearchDataString()
		dim dataString
		dataString = "<ReportViewData UID=""" & me.Name & """>"
					
		'open fields
		dataString = dataString & "<Fields>"					
		'loop the fields
		dim field
		for each field in me.Fields
			dataString = dataString & "<Field name=""" & field & """/>" 
		next
		'close fields
		dataString = dataString & "</Fields>"
		'open rows
		dataString = dataString & "<Rows>"
		dim result, resultField, i
		for each result in me.Results
			'open row
			dataString = dataString & "<Row>"
			for i = 0 to result.Count -1
				resultField = result(i)
				field = m_Fields(i)
				dataString = dataString & "<Field name=""" & field & """ value=""" & resultField & """/>" 
			next
			'close row
			dataString = dataString & "</Row>"
		next
		'close rows
		dataString = dataString & "</Rows>"
		'close ReportViewData
		dataString = dataString & "</ReportViewData>"
		'return data
		makeSearchDataString = dataString
	end function
	
'#endregion functions	
End class