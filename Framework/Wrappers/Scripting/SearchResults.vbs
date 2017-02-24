'[path=\Framework\Wrappers\Scripting]
'[group=Wrappers]

!INC Utils.Include

'Author: Geert Bellekens
'Date: 2015-12-07
'dim outputTabName
'outputTabName = "ModelValidation"

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
		dim dataString
'		Repository.WriteOutput outputTabName, now() & " starting makeDataString",0
		dataString = makeSearchDataString()
'		Repository.WriteOutput outputTabName, now() & " Datastring: " & dataString,0
'		Repository.WriteOutput outputTabName, now() & " finished makeDataString",0
		Repository.RunModelSearch "","","", dataString
	end function
	
	private function makeSearchDataString()
		
		dim xmlDOM 
		set  xmlDOM = CreateObject( "Microsoft.XMLDOM" )
		'set  xmlDOM = CreateObject( "MSXML2.DOMDocument.4.0" )
		xmlDOM.validateOnParse = false
		xmlDOM.async = false
		 
		dim xmlRoot 
		set xmlRoot = xmlDOM.createElement( "ReportViewData" )
		dim uidAttr 
		set uidAttr = xmlDOM.createAttribute("UID")
		uidAttr.nodeValue = me.Name
		xmlRoot.setAttributeNode(uidAttr)
		xmlDOM.appendChild xmlRoot

		dim xmlFields
		set xmlFields = xmlDOM.createElement( "Fields" )
		xmlRoot.appendChild xmlFields
		'loop the fields
		dim field
		for each field in me.Fields
			dim xmlField 
			set xmlField = xmlDOM.createElement( "Field" )
			dim nameAttr
			set nameAttr = xmlDOM.createAttribute("name")
			nameAttr.nodeValue = field
			xmlField.setAttributeNode(nameAttr)
			xmlFields.appendChild xmlField
		next
		'add rows
		dim xmlRows
		set xmlRows = xmlDOM.createElement( "Rows" )
		xmlRoot.appendChild xmlRows
		'add row
		dim result, resultField, i
		for each result in me.Results
			dim xmlRow
			set xmlRow = xmlDOM.createElement( "Row" )
			xmlRows.appendChild xmlRow
			'add fields
			for i = 0 to result.Count -1
				resultField = result(i)
				field = m_Fields(i)
				'field attribute
				set xmlField = xmlDOM.createElement( "Field" )
				set nameAttr = xmlDOM.createAttribute("name")
				nameAttr.nodeValue = field
				'value attribute
				xmlField.setAttributeNode(nameAttr)
				dim valueAttr
				set valueAttr = xmlDOM.createAttribute("value")
				valueAttr.nodeValue = resultField
				xmlField.setAttributeNode(valueAttr)
				'add the field to the row
				xmlRow.appendChild xmlField
			next
		next
		'return
		makeSearchDataString = xmlDOM.xml
		
	end function
	
'#endregion functions	
End class