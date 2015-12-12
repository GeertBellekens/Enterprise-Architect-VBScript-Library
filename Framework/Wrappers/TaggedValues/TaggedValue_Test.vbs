'[path=\Framework\Wrappers\TaggedValues]
'[group=Testing]
Option Explicit

!INC Local Scripts.EAConstants-VBScript
!INC TaggedValues.TaggedValue

''' --------------------------
''' TESTING TAGGEDVALUE HELPER
''' --------------------------

Private Sub Module_Initialize()
	Set m_tagapi = Nothing
End Sub

Private Sub Module_Terminate()
	Set m_tagapi = Nothing
End Sub


Sub Test_ExtractMethod()
	Dim sTestString
	sTestString = _
		"VBA.Persistence : Type=Enum;" & Chr(10) & _
		"Values=Persistent;Transient;" & Chr(10) & _
		"Default=Persistent" & Chr(10) & _
		"BaseStereotype=class;attribute;... etc.;" & Chr(10)

	''' Testing the robustness of the extract method, test data 
	''' lacks of a proper eol-char:
	''' "Default=Persistent"    <- no trailing semicolon;
	
	''' Run the test	
	FormatPropertyData sTestString
	
	MsgBox ExtractStrValueFromRawdata(sTestString, "BaseStereotype")
	
	MsgBox ExtractStrValueFromRawdata(sTestString, "Type") & Chr(10) & _
			ExtractStrValueFromRawdata(sTestString, "Values") & Chr(10) & _
			ExtractStrValueFromRawdata(sTestString, "Default") & Chr(10) & _
			ExtractStrValueFromRawdata(sTestString, "BaseStereotype")
End Sub


Private Sub TEST_PrintTestCase(tv, aDoPrintProperties)
	Dim S
	Dim t As EA.TaggedValue
	Set t = tv
	If aDoPrintProperties Then
		TagHelp.Wrap(t)
		Session.Output "Name         : " & m_tagapi.Name()
		Session.Output "Value        : " & m_tagapi.Value()
		Session.Output "FQName       : " & m_tagapi.FQName()
		Session.Output "HasNotes     : " & m_tagapi.HasNotes()
		Session.Output "Notes        : " & m_tagapi.Notes()
		Session.Output "HasMemo      : " & m_tagapi.HasMemo()
		
		Session.Output "PropertyID   : " & m_tagapi.PropertyID()
		Session.Output "PropertyGUID : " & m_tagapi.PropertyGUID()		
		
		Session.Output "M_Value      : " & m_tagapi.M_Value()
		Session.Output "M_Default    : " & m_tagapi.M_Default()
		Session.Output "M_GlobalDefault:"& m_tagapi.M_GlobalDefault()
		
		Session.Output "ParentName  : " & m_tagapi.ParentName()
		Session.Output "ParentType  : " & m_tagapi.ParentType()
		Session.Output "ParentID    : " & m_tagapi.ParentID()
		If m_tagapi.Wrap(t).IsRoleTag Then
			Session.Output "ParentObject.Name: " & m_tagapi.ParentObject.Role()
		Else
			Session.Output "ParentObject.Name: " & m_tagapi.ParentObject.Name()
		End If
		
		Session.Output "IsPackageTag   : " & m_tagapi.IsPackageTag
		Session.Output "IsClassTag     : " & m_tagapi.IsClassTag
		Session.Output "IsTaggedValue  : " & m_tagapi.IsTaggedValue
		Session.Output "IsElementTag   : " & m_tagapi.IsElementTag
		Session.Output "IsInterfaceTag : " & m_tagapi.IsInterfaceTag
		Session.Output "IsAttributeTag : " & m_tagapi.IsAttributeTag
		Session.Output "IsMethodTag    : " & m_tagapi.IsMethodTag
		Session.Output "IsConnectionTag: " & m_tagapi.IsConnectionTag
		Session.Output "IsRoleTag      : " & m_tagapi.IsRoleTag
		Session.Output " "
	Else
		If TagHelp.Wrap(t).TryValue(S) then
			S = S & TagHelp.Name()
		Else
			S = S & TagHelp.Name()
		End If

		S = S & m_tagapi.PropertyID()
		S = S & m_tagapi.PropertyGUID()
		
		S = S & m_tagapi.HasNotes()
		S = S & m_tagapi.Notes()
		S = S & m_tagapi.HasMemo()
		
		S = S & m_tagapi.FQName()
		
		S = S & m_tagapi.M_Value()
		S = S & m_tagapi.M_Default()
		S = S & m_tagapi.M_GlobalDefault()
		
		S = S & m_tagapi.ParentName()
		S = S & m_tagapi.ParentType()
		S = S & m_tagapi.ParentID()
		
		Select Case m_tagapi.M_ParentObjectType
		    Case EA_Attribute
					Dim attr As EA.Attribute
					Set attr = m_tagapi.M_ParentObject
					S = S & attr.Name()
		    Case EA_Method
					Dim meth As EA.Method
					Set meth = m_tagapi.M_ParentObject
					S = S & meth.Name()
		    Case EA_Connector
					Dim conn As EA.Connector
					Set conn = m_tagapi.M_ParentObject
					S = S & conn.Name()
			Case EA_ConnectorEnd
					Dim role As EA.ConnectorEnd
					Set role = m_tagapi.M_ParentObject
					S = S & role.Role()
			Case EA_Class
					Dim elem As EA.Element
					Set elem = m_tagapi.M_ParentObject
					S = S & elem.Name()
			Case EA_Package
					Dim pack As EA.Package
					Set pack = m_tagapi.M_ParentObject
					S = S & pack.Name()
			Case Else
					Err.Raise err_ElementType, msg_ElementType '// Indicate error					
		End Select
		
		S = S & m_tagapi.IsAttributeTag
		S = S & m_tagapi.IsMethodTag
		S = S & m_tagapi.IsConnectionTag
		S = S & m_tagapi.IsRoleTag
		S = S & m_tagapi.IsTaggedValue
		S = S & m_tagapi.IsClassTag
		S = S & m_tagapi.IsElementTag
		S = S & m_tagapi.IsInterfaceTag
		S = S & m_tagapi.IsPackageTag	
		S = S & ""
	End If
End Sub

Private Function TagHelp()
	if m_tagapi is Nothing then _
		Set m_tagapi = New TTaggedValueWrapper
	Set TagHelp = m_tagapi
End Function

Private Sub Test_ListAllTaggedValuesForSelectedPackage(DoPrintProperties)

	Dim p As EA.PropertyType
	
	Dim et As EA.TaggedValue ''' Same for Package and Class (Element)
	Dim Package as EA.Package
	Set Package = Repository.GetTreeSelectedPackage()
	
	''' Print repository
	If DoPrintProperties Then
		Session.Output "-------------------------------------------------"
		Session.Output " All PropertyTypes in internal respository       "
		Session.Output "-------------------------------------------------"	
	End If
	
	Dim tmp as EA.Element
	Dim e as EA.Element

	TagHelp.StatsStart
	
	''' PACKAGES
	If DoPrintProperties then _
		if Package.Packages.Count>0 then Session.Output "** PACKAGE ********* "
	For Each et in Package.Element.TaggedValues
		TEST_PrintTestCase et, DoPrintProperties
	Next
	
	Dim pkg as EA.Package
	For Each pkg in Package.Packages
		For Each et in pkg.Element.TaggedValues
			TEST_PrintTestCase et, DoPrintProperties
		Next
	Next

	For Each e in Package.Elements
	
		Session.Output "Name         : "  & e.Name()
		If m_tagapi.TryWrapByName("VBA.FileName", e)  then
			TEST_PrintTestCase m_tagapi.tvObject, DoPrintProperties
		End If
		
		''' CLASSES
		If DoPrintProperties then Session.Output "= CLASS ======"	
		For Each et in e.TaggedValues
			TEST_PrintTestCase et, DoPrintProperties
		Next

		''' ATTRIBUTES
		
		If DoPrintProperties then _
			If e.Attributes.Count>0 then Session.Output "- Attributes"
		Dim a As EA.Attribute
		Dim at As EA.AttributeTag
		
		For Each a in e.Attributes
			For Each at in a.TaggedValues
				TEST_PrintTestCase at, DoPrintProperties
			Next
		Next

		''' METHODS / OPERATIONS
		If DoPrintProperties then _
			If e.Methods.Count>0 then Session.Output "- Methods"
		Dim m As EA.Method
		Dim mt As EA.MethodTag
		
		For Each m in e.Methods
			For Each mt in m.TaggedValues
				TEST_PrintTestCase mt, DoPrintProperties
			Next
		Next

		''' CONNECTORS
		If DoPrintProperties then _
			If e.Connectors.Count>0 then Session.Output "- Connectors"
			
		Dim c As EA.Connector
		Dim ct As EA.ConnectorTag
		Dim r as EA.ConnectorEnd
		Dim rt as EA.RoleTag
		
		For Each c in e.Connectors
			If DoPrintProperties then Session.Output " - Connector -"
			For Each ct in c.TaggedValues
				TEST_PrintTestCase ct, DoPrintProperties
			Next
				Dim sGuid
				Dim conn as EA.Connector
				
				''' ROLES
				Set r = c.ClientEnd
				For Each rt in r.TaggedValues
					If DoPrintProperties then Session.Output "-- Client Role --"
					TEST_PrintTestCase rt, DoPrintProperties
				Next
				
				Set r = c.SupplierEnd
				For Each rt in r.TaggedValues
					If DoPrintProperties then Session.Output "-- Supplier Role --"
					TEST_PrintTestCase rt, DoPrintProperties
				Next
		Next
		
		''' Pausing Stats will accumulate ("Acc") the time since last 
		''' pause-resume. StatsStart will reset all counters
		TagHelp.StatsPause
	'		Session.Output "Hits (Wraps)        : " & m_tagapi.StatsWrapCount()
	'		Session.Output "Hits                : " & m_tagapi.StatsCount()
	'		Session.Output "Hits (acc)          : " & m_tagapi.StatsCountAcc()
			Session.Output "Duration            : " & Minute( m_tagapi.StatsDuration()) & ":" & Round( Second(m_tagapi.StatsDuration()), 3)
			Session.Output "Duration       (Acc): " & Minute( m_tagapi.StatsDurationAcc()) & ":" & Round( Second(m_tagapi.StatsDurationAcc()), 3)
	'		Session.Output " "
		TagHelp.StatsResume
	Next

	TagHelp.StatsStop
	if TagHelp.HasStats then
		Session.Output "Statistics"
		Session.Output "-----------------------------------------------"
		Session.Output "Hits         (Wraps): " & m_tagapi.StatsWrapCount()
		Session.Output "Hits                : " & m_tagapi.StatsCount()
		Session.Output "Hits           (Acc): " & m_tagapi.StatsCountAcc()
		Session.Output "Hits Per Second     : " & Round( m_tagapi.StatsHitsPerSecond(), 3)
		Session.Output "Hits Per Second(Acc): " & Round( m_tagapi.StatsHitsPerSecondAcc(), 3)
		Session.Output "-----------------------------------------------"
		Session.Output "Duration            : " & Minute( m_tagapi.StatsDuration()) & ":" & Round( Second(m_tagapi.StatsDuration()), 3)
		Session.Output "Duration       (Acc): " & Minute( m_tagapi.StatsDurationAcc()) & ":" & Round( Second(m_tagapi.StatsDurationAcc()), 3)
		Session.Output "Time Per Hit        : " & Round( Second(m_tagapi.StatsTimePerHits()), 3) & " sec"
		Session.Output "Time Per Hit   (Acc): " & Round( Second(m_tagapi.StatsTimePerHitsAcc()), 3) & " sec"
		Session.Output "-----------------------------------------------"
		Session.Output " "
	End If
	Set m_tagapi = Nothing
End Sub



''' MAIN

Sub Main()
	Dim DoPrintProperties
	
	Repository.EnsureOutputVisible( "Script" )
	
	Session.Output "--++:::: START: " & Date() & " -- " & Time() & " ::::++--"
	
	Module_Initialize()
	'Test_ExtractMethod()
	'DoPrintProperties = False
	DoPrintProperties = True
	Test_ListAllTaggedValuesForSelectedPackage(DoPrintProperties)
'	DoPrintProperties = True
'	Test_ListAllTaggedValuesForSelectedPackage(DoPrintProperties)
	
	Session.Output "--++:::: STOP: " & Date() & " -- " & Time() & " ::::++--"	
	Module_Terminate()	
End Sub

Main