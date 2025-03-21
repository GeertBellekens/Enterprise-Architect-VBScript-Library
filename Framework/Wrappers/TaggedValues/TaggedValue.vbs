'[path=\Framework\Wrappers\TaggedValues]
'[group=Wrappers]
'Option Explicit

!INC Local Scripts.EAConstants-VBScript

''' ===========================================================================
''' TAGGEDVALUE HELPER
''' ===========================================================================
''' VERSION			: 0.9.6
'''	RELEASE DATE	: 2015-12-10
''' HISTORY			: See History.txt				First release in 2015-12-07
'''
''' DESCRIPTION		: A TaggedValue Helper wrapper, intending to provide access 
'''					  to TaggedValue properties with consistent (orthogonal)  
''' 				  property names. More info far below.
''' 
''' AUTHOR			: Rolf Lampa, RIL Partner AB, rolf.lampa@rilnet.com
'''
''' COPYRIGHT		: (C) Rolf Lampa, 2015. Free to use for commercial projects 
'''				  	  if giving proper attribution to the author and providing 
'''					  this copyright info visible in your code and product 
'''					  documentation, including donation info below.
'''
''' DONATIONS		: If you find the script being useful you may consider 
'''					  making a donation. All amounts amounts. For Paypal 
'''					  donations, use the following url:
'''
''' https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=KJCD6N8M8MRWQ
'''
''' DEPENDENCIES 	: None. The script should work as is inside Enterprise 
'''					  Architect. 
''' TESTED			: Run on Enterprise Architect 12.1 Beta, using the file
'''					  "TaggedValue_Test.vbs" for simple property access
'''
''' ---------------------------------------------------------------------------
'''
''' USAGE: The code below (for example) will provide a "singleton" access to 
'''		   the TaggedValue helper:

Dim m_tagapi 	''' You may want to use this variable directly, the after first 
				''' access to the TagAPI

Public Function TagAPI()
	''' Ensure that the TaggedVaslue helper is created only once.
	if m_tagapi is Nothing then
		Set m_tagapi = New TTaggedValueWrapper
	End If
	Set TagAPI = m_tagapi
End Function

Private Sub Module_Initialize()
	''' Ensure initialization of the variable as to prepare 
	''' for the assigment check in the TagAPI() function
	Set m_tagapi = Nothing
End Sub

Private Sub Module_Terminate()
	''' Call this explicitly to dispose of the object.
	Set m_tagapi = Nothing
End Sub

''' ---------------------------------------------------------------------------
''' EXAMPLE OF USAGE:
'''
'''		Dim tv As EA.TaggedValue
'''		For Each tv in Pkg.Element.TaggedValues
''' 		Session.Output TagAPI.Wrap(tv)				''' Wrap...
''' 		Session.Output TagAPI.Wrap(tv).Value()		''' or Wrap and use directly.
''' 		Session.Output TagAPI.Name()				''' Now use the helper Obj directly
''' 		Session.Output m_tagapi.Notes()				''' Or use the module variable
''' 		Session.Output m_tagapi.FQName()					
''' 		Session.Output m_tagapi.PropertyGUID()					
''' 		Session.Output m_tagapi.ParentName()					
''' 		Session.Output m_tagapi.ParentID()					
''' 		Session.Output m_tagapi.ParentGUID()			
''' 		''' Etc
''' 	Next
''' 
''' 	One can also use a class, and set the wrapper "ByName", then the wrapper
''' 	looks up the TaggedValue, wraps it, and exposes its properties, like so:
'''
'''		For Each elem in Pkg.Elements
'''			If TagApi.TryWrapByName("VBA.FileName", elem)  then
''' 			Session.Output m_tagapi.Name()			
''' 			Session.Output m_tagapi.Notes()			
''' 			Session.Output m_tagapi.PropertyGUID()
''''		End If
'''		Next
'''
'''		''' Or, if assuming that the TV exists:
'''
''' 	S = TagApi.WrapByName("VBA.FileName", elem).Notes()
'''		
''' ---------------------------------------------------------------------------
'''
''' STATISTICS
''' 
''' Simple statistics is supported but can be "disabled" from the code altogether 
''' by using the following Regex Expression:
'''
''' DISABLE all stats code (comment):
''' 	Regex Search:	^(?!'//)(.*?\(\(\$stats\)\).*?$)
'''		Regex Replace:	'//\1
''' ENABLE stats code again (uncomment):
''' 	Regex Search:	^(?='//)'//(.*?\(\(\$stats\)\).*?$)
'''		Regex Replace:	\1

''' ---------------------------------------------------------------------------
''' TODO:
''' - Property IsInterfaceTag() - Check Stereotype to distinguish from regular 
'''   Class
'''
''' ---------------------------------------------------------------------------
''' CLASS MEMEBERS
''' ---------------------------------------------------------------------------
'''
'''		''' Most used properties & functions :
'''
'''     ''' WrapByName: Direct access to properties assuming the TV exists
'''		Public Function WrapByName(aName, ByRef aObj) ''': TTaggedValueWrapper
'''		Public Function TryWrapByName(aName, ByRef aObj) ''': Boolean
'''		Public Function Wrap(ByRef aTaggedValue) ''': TTaggedValueWrapper
'''		Public Property Get Value() ''': String
'''		Public Function TryValue(ByRef S) ''': Boolean
'''		Public Property Get Name()	''': String
'''		Public Property Get Notes()	''': String
'''		Public Sub Update() ''': Void			''' All PropertyTypes reloaded from Repository. Total reinitialization
'''		
'''		'// Statistics control :
'''		
'''		Public Sub StatsPause() ''': Void		''' Collects Duration and HitCounts into ditto "Accumulated"
'''		Public Sub StatsResume() ''': Void		''' Restarts with HitCount=0 (AccumulatedCount continues though)
'''		Public Sub StatsStart() ''': Void		''' Resets all counters and timers
'''		Public Sub StatsStop() ''': Void		''' Keeps all count and time info until next StatsStart()

'''		''' Other useful and orthogonal properties.
'''		''' In case a value isn't actually provided by the underlaying object, 
'''		''' these properties at least provides with a fake value as to allowing 
'''		''' for "type safe" when traversing EA models.
'''		
'''		Public Property Get Detail() ''': String
'''		Public Property Get FQName() ''': String
'''		Public Property Get HasMemo() ''': Boolean
'''		Public Property Get HasNotes()	''': String
'''		Public Property Get HasStats() ''': Boolean
'''		Public Property Get IsAttributeTag() ''': Boolean
'''		Public Property Get IsClassTag() ''': Boolean
'''		Public Property Get IsConnectionTag() ''': Boolean
'''		Public Property Get IsElementTag() ''': Boolean
'''		Public Property Get IsInterfaceTag() ''': Boolean
'''		Public Property Get IsMethodTag() ''': Boolean
'''		Public Property Get IsPackageTag() ''': Boolean
'''		Public Property Get IsRoleTag() ''': Boolean
'''		Public Property Get IsTaggedValue() ''': Boolean
'''		Public Property Get IsValueDefault() ''': Boolean
'''		Public Property Get M_Default()	''': String
'''		Public Property Get M_GlobalDefault()	''': String
'''		Public Property Get M_IsRoleTag() ''': Boolean
'''		Public Property Get M_ParentObject() ''': EA.<Object>
'''		Public Property Get M_Value() ''': String
'''		Public Property Get TvObject() ''': EA.TaggedValue		''' Useful with WrapByName
'''		Public Property Get ParentID() ''': Integer
'''		Public Property Get ParentName() ''': String
'''		Public Property Get ParentObject() ''': EA.<Object>
'''		Public Property Get ParentObjectType() ''': Integer (ot<ObjectType>)
'''		Public Property Get ParentType() ''': String (Kind name)
'''		Public Property Get PropertyGUID() ''': String
'''		Public Property Get PropertyID() ''': String
'''		Public Property Get StatsCount() ''': Integer
'''		Public Property Get StatsCountAcc() ''': Integer
'''		Public Property Get StatsDuration() ''': Time
'''		Public Property Get StatsDurationAcc() ''': Time
'''		Public Property Get StatsHitsPerSecond() ''': Integer
'''		Public Property Get StatsHitsPerSecondAcc() ''': Integer
'''		Public Property Get StatsTimePerHits() ''': Time
'''		Public Property Get StatsTimePerHitsAcc() ''': Integer
'''		Public Property Get StatsWrapCount() ''': Integer
'''		
'''		Private Function ContainsStr(aStr, aChar)
'''		Private Function ExtractPropertyFromRawStr(ByRef aSubjectStr, ByVal aFieldName) ''': Boolean
'''		Private Function GetValueByXmlTagName(ByRef aStr, ByRef aTagName, ByRef OutResult) ''': String, Boolean
'''		Private function PropertyTypeByName(aNameAsKey, ByRef OutProp) ''': PropertyType, Boolean
'''		Private Function QueryRoleTagForElementID(ByRef OutGUID) ''': Boolean
'''		Private Function TryExtractRoleTagStereotypeDefault(ByRef S) ''': String, Boolean
'''		Private Function TryExtractRoleTagValue(ByRef S) ''': Boolean
'''		Private Function TryExtractStereotypeDefault(ByRef s) ''': Boolean		
'''		Private function TryGetPropertyTypeDefault(aNameAsKey, ByRef OutResult) ''': String, Boolean
'''		Private Property Get ConnectionEndForRoleTag()
'''		Private Property Get ConnectorForRoleTag()
'''		Private Property Get IsClient() ''': Boolean
'''		Private Property Get IsSource() ''': Boolean
'''		Private Property Get IsSupplier() ''': Boolean
'''		Private Property Get IsTarget() ''': Boolean
'''		Private Property Get PropertyTypesDefaultDictionary()  ''': Dictionary
'''		Private Property Get PropertyTypesDictionary
'''		Private Property Get PropertyTypesRawDataDictionary() ''': Dictionary
'''		Private Property Get RoleTagConnector()
'''		Private Property Let UseStats(aBool) ''': Void	''' (($stats))
'''		
'''		Private Sub Class_Initialize() ''': Void
'''		Private Sub Class_Terminate() ''': Void
'''		
'''		Private Sub FormatPropertyTypesText(ByRef aSubjectStr)
'''		Private Sub IncStats() ''': Void					''' (($stats))
'''		Private Sub LoadPropertyData()
'''		Private Sub RegisterPropertyTypes() ''': Void
'''		Private Sub RegisterPropertyTypesDefaults() ''': Void
'''		Private Sub RegisterPropertyTypesRawData() ''': Void
'''		Private Sub ResetData() ''': Void
'''		Private Sub ResetStats() ''': Void					''' (($stats))
''' -----------------------------------------------------------------------


''' TAGGEDVALUE TYPES

Public Const EA_TaggedValue 	= 12	''' Classes & Packages
Public Const EA_AttributeTag	= 34	''' Attributes
Public Const EA_MethodTag		= 36	''' Methods
Public Const EA_ConnectorTag	= 38	''' Connectors
Public Const EA_RoleTag			= 41	''' Role/ConnectorEnd

''' MODEL ELEMENTS

Public Const EA_Element 		= 4		''' Class & Interface (see <stereotype>!)
Public Const EA_Class 			= 4		''' -"-
Public Const EA_Interface		= 4		''' -"-
Public Const EA_Package			= 5		''' Package
Public Const EA_Attribute		= 23	''' Attributes
Public Const EA_Method			= 24	''' Methods
Public Const EA_Connector		= 7		''' Connectors
Public Const EA_Role			= 22	''' Role/ConnectorEnd
Public Const EA_ConnectorEnd	= 22	''' Role/ConnectorEnd

Public Const EA_ASSOCIATION_SOURCE = "ASSOCIATION_SOURCE"
Public Const EA_ASSOCIATION_TARGET = "ASSOCIATION_TARGET"

''' ERROR CODES & MESSAGES

Dim err_ElementType : err_ElementType = vbObjectError + 1
Dim err_TaggedValueType : err_TaggedValueType = vbObjectError + 1

Private Const msg_ElementType = "Invalid Element Type!"
Private Const msg_TaggedValueType = "Invalid TaggedValue type!"

''' HELPER CLASS

Class TTaggedValueWrapper
	
	''' The currently wrapped EA.TaggedValue + EA.RoleTag.
	Dim m_tv As EA.TaggedValue
	
	''' EA.RoleTag (special case) introduced for optmization reasons. It is 
	'''	being assigned already at the time of Wrap(tv)
	Dim m_rt As EA.RoleTag
	
	''' The owning EA.Connector (only) for visiting EA.RoleTags.
	Dim m_roletag_connector As EA.Connector
	
	''' Optimized Access, Lazy evaluation in ParentObject (but or course, needs a type cast for usage)
	Dim m_parentobj As EA.Element
	Dim m_parent_typename						'// For chaching purpose.
	Dim m_parent_objecttype
	
	Dim m_objecttype
	
	''' Used by InStr() & Mid() when extracting values from text. Avoids repeated allocs.
	Dim m_startpos
	Dim m_endpos
	
	''' Dictionary for fast RT access of Property OBJECTS
	''' accessed by TaggedValue name
	Dim m_PropertyTypesDictionary
	
	''' Dictionary for fast direct RT access of property Defaults, 
	''' in String format, accessed directly by TaggedValue Name()
	Dim m_PropertyTypesDefaultDictionary
	
	''' Dictionary for storing property data from .Detail property, 
	''' in a prepared format making it easier (faster) to extract any
	''' individual property value from its multiline text content. 
	''' Performance and the utilization of Lazy Evaluation motivates
	''' these extra dictionaries.	
	Dim m_PropertyTypesRawDataDictionary

	''' Used by WrapByName to lookup and wrap a TaggedValue by name
	Dim m_elem As EA.Element

	''' STATS                               ''' Keep the "(($stats))" markers (used for Regex enable/disable)
	
	Dim m_usestats							''' (($stats)) Set to false in production code!
	Dim m_stats_hitcount					''' (($stats)) Counts all essesacc to the properties
	Dim m_stats_hitcount_acc				''' (($stats))
	Dim m_stats_wrapcount					''' (($stats)) Counts all assignments with Wrap method since created
	Dim m_stats_starttime					''' (($stats))
	Dim m_stats_stoptime					''' (($stats))
	Dim m_stats_time_acc					''' (($stats))
	
	''' INITIALIZERS
	
	Private Sub Class_Initialize() ''': Void
		ResetStats()						''' (($stats))
		m_usestats = False					''' (($stats)) Must be set explicitly
		Update()
	End Sub
	
	
	Private Sub Class_Terminate() ''': Void
		ResetData()
		m_usestats = False					''' (($stats)) Must be set explicitly
	End Sub
	
	
	''' [TryWrapByName]
	''' Provide a class, an attribute, a connector etc, and a name for the TaggedValue.
	''' If success (result true) then the properties of the TV is immediately available, 
	''' like so: 
	''' 	If TagHelp.TryWrapByName("<TaggedValueName>", Pkg) then 
	''' 		S = TagHelp.Value() ''' etc
	''' 	End If
	Public Function TryWrapByName(aName, ByRef aObj) ''': Boolean
		Dim was_found
		Dim tv As EA.TaggedValue
		Set m_elem = aObj
		Select Case m_elem.ObjectType
			Case EA_Element
				''' m_elem was already set on entry, so go ahead
				Set tv = m_elem.TaggedValues.GetByName(aName)
			Case EA_Package	
				Dim m_pack As EA.Package
				Set m_pack = aObj
				Set tv = m_pack.Element.TaggedValues.GetByName(aName)
			Case EA_Attribute
				Dim m_attr As EA.Attribute
				Set m_attr = aObj
				Set tv = m_attr.TaggedValues.GetByName(aName)
			Case EA_Method
				Dim m_meth As EA.Method
				Set m_meth = aObj
				Set tv = m_meth.TaggedValues.GetByName(aName)
			Case EA_Connector
				Dim m_conn As EA.Connector
				Set m_conn = aObj
				Set tv = m_conn.TaggedValues.GetByName(aName)
			Case EA_ConnectorEnd
				Dim m_role As EA.ConnectorEnd
				Set m_role = aObj
				Set tv = m_role.TaggedValues.GetByName(aName)
			Case Else
				Err.Raise err_ElementType, msg_ElementType
		End Select
		was_found = Not tv is Nothing
		If was_found then
			TryWrapByName = Not Wrap(tv).TvObject Is Nothing
		Else
			TryWrapByName = False
		End If
	End Function
	
	
	''' [WrapByName]
	''' Use the result (object) directly when you know for certain that the 
	''' named TaggedValue will be found, like so:
	'''	S = TagAPI.TagByName("SomeTagName", elem).Value()      or,
	'''	S = TagAPI.TagByName("CopyrightNotice", elem).Notes()  etc.
	Public Function WrapByName(aName, ByRef aObj) ''': TTaggedValueWrapper
		If TryWrapByName(aName, aObj) then
			Set WrapByName = Me
		Else
			Set WrapByName = Nothing
		End If
	End Function
	
	
	''' [Wrap] 
	''' Assigns the external TaggedValue to the wrapper class 
	''' for use in the internal processing.
	Public Function Wrap(ByRef aTaggedValue) ''': TTaggedValueWrapper
		''' m_tv is most often used, if not a RoleTag arrives (then m_rt instead)
		Set m_tv = aTaggedValue
		m_objecttype = m_tv.ObjectType
		''' Cast if RoleTag, this extra property of RoleTag tyoe can now be both 
		''' tested for ("is Nothing") and used directly by any internal properties.
		If m_objecttype = EA_RoleTag Then
			Set m_rt = m_tv
		Else
			Set m_rt = Nothing
		End If
		
		''' This is TaggedValue.ParentObj.ObjectType
		m_parent_objecttype	= 0
		m_startpos			= 0
		m_endpos			= 0
		''' Used internally as a ' cache', sorts of
		m_parent_typename = ""
		''' Lazy Evaluation in Property Get RoleTagConnector													
		Set m_parentobj = Nothing
		Set m_roletag_connector = Nothing
		
		If m_usestats Then 									''' (($stats))
			m_stats_wrapcount = m_stats_wrapcount + 1		''' (($stats))
		End If												''' (($stats))
		
		''' Return this wrapped object as to provide immediate 
		''' access to the wrapper's functionality, like so:
		'''
		''' 	Set tvapi = New TRILTaggedValueApi
		''' 	S = tvapi.Wrap(aTV).Value(), then
		'''		S = tvapi.Value()							''' Avoid using the wrap function more than once
		
		Set Wrap = Me										''' Returns the wrapper itself.
	End Function
	
	
	''' ----------------------------------------------------------
	''' PUBLIC PROPERTIES
	''' ----------------------------------------------------------
	
	
	''' [Value]
	''' Derives any Default value if the tv Value() is empty. Use TryValue(S) 
	''' if you want to utilize a Boolean reply whether any value  at all was 
	''' returned.
	''' This property was the main reason why this wrapper was designed in 
	''' the first place. It attempts to return Value(), and if not exists, 
	''' it tries to get a Default (from "initial value" in Stereotypes), 
	''' and if no value was found there either it reads from GlobalDefault, 
	'''	which has its default values defined in Repository.PropertyTypes() 
	''' and stored in common table 't_propertytypes'.
	''' If you want "direct values" without any of the below semantics applied, 
	''' call M_Value or M_Default, or M_GlobalDefault directly.
	Public Property Get Value() ''': String
		Dim S
		IncStats()										''' (($stats))
		If TryValue(S) Then
			Value = S
		Else
			Value = ""
		End If
	End Property
	
	
	''' [TryValue]
	''' Se documentation for Property Value(). See also EA's documentation page
	''' on how the value content is to be interpreted (rules which we "hide" and just 
	''' deliver in this wrapper) :
	''' http://sparxsystems.com/enterprise_architect_user_guide/12.1/automation_and_scripting/taggedvalue.html
	Public Function TryValue(ByRef S) ''': String
		IncStats()										''' (($stats))
		S = m_tv.Value
		
		''' SPECIAL CASE (RoleTag)
		If m_objecttype = EA_RoleTag Then
			
			''' EA.RoleTag type: Strip out value before "$ea_notes="
			''' The RoleTag must be processed entirely in this block (thus the ElseIf) 
			''' since it's not typecompatible with the other properties for retriveing
			''' Default values.
			
			If TryExtractRoleTagValue(S) Then
				TryValue = True
			ElseIf TryExtractStereotypeDefault(S) Then
				TryValue = True
			ElseIf TryGetPropertyTypeDefault(m_rt.Tag, S) Then	''' (Params = Name, S)				
				''' OK to use common method also for this (RoleTag) type, since
				''' global Defaults are not store in the (Role) TaggedValue itself,
				''' but in the Repository.PropertyTypes, which are common for all.									
				TryValue = True
			Else
				S = ""
				TryValue = False
			End If
			''' NORMAL CASE(S)
		ElseIf  S = "" Then
			''' Try retrive Default ("Initial value" from Stereotype:
			If TryExtractStereotypeDefault(S) Then			
				TryValue = True
				
			ElseIf TryGetPropertyTypeDefault(m_tv.Name, S) Then
				''' As a Last resort, try Default from t_PropertyTypes if
				''' it (this Global) was not "overrided" in a Stereotype.
				TryValue = True
			Else
				''' Neither a Value nor a Default value was found.
				TryValue = False
				S = ""
			End If
		ElseIf S = "<memo>" Then
			''' If "<memo>" the value shall be retrived from .Notes
			S = m_tv.Notes
			TryValue = True
		Else
			TryValue = True 			''' S already contains the value
		End If
	End Function
	
	
	''' [Name]
	''' Special case RoleTag which stores "Name" in its "Tag" property.
	Public Property Get Name()	''': String
		IncStats()									''' (($stats))
		If m_objecttype = EA_RoleTag Then
			Name = m_rt.Tag()
		Else 
			''' In all other cases; EA_TaggedValue, EA_AttributeTag, EA_MethodTag, EA_ConnectorTag
			Name = m_tv.Name()
		End If
	End Property
	
	
	''' [Notes]
	Public Property Get Notes()	''': String
		IncStats()									''' (($stats))
		If m_objecttype = EA_RoleTag Then
			Notes = ""			''' RoleTags doesn't have any notes (Only ConnectorEnd has)
		Else
			Notes = m_tv.Notes
		End If
	End Property
	
	
	''' [PropertyGUID]
	''' Different property names ("TagGUID" and "PropertyGUID") for different Tag owners
	''' requires type check before accessing the properties without crashing :
	Public Property Get PropertyGUID() ''': String
		Select Case m_objecttype
			Case EA_TaggedValue
				' Direct use of the internal ref 'm_tv'
				PropertyGUID = m_tv.PropertyGUID
			Case EA_AttributeTag	
				Dim at As EA.AttributeTag
				Set at = m_tv
				PropertyGUID = at.TagGUID
			Case EA_MethodTag
				Dim mt As EA.MethodTag
				Set mt = m_tv
				PropertyGUID = mt.TagGUID
			Case EA_ConnectorTag	
				Dim ct As EA.ConnectorTag
				Set ct = m_tv
				PropertyGUID = ct.TagGUID
			Case EA_RoleTag
				''' No property ID exist for RoleTag! But we provide a fake ID anyway. 
				''' In any case the user must check the result when calling GetTaggedValueByID(Id)
				PropertyGUID = m_rt.PropertyGUID
			Case Else
				Err.Raise err_TaggedValueType, msg_TaggedValueType
		End Select				
	End Property
	
	
	''' [PropertyID]
	''' Missing property for Roletag. Returns only a Fake ID (-1, since it doesn't have any 
	''' integer ID at all, only a PropertyGUID), but we still provide a value in order to 
	''' avoid type errors for users which are looping through model info.	
	''' Different Tag owners all have different property names.
	Public Property Get PropertyID() ''': String
	Select Case m_objecttype
		Case EA_TaggedValue
    		' Direct use of the internal ref 'm_tv'
    		PropertyID = m_tv.PropertyID
		Case EA_AttributeTag	
    		Dim at As EA.AttributeTag
    		Set at = m_tv
    		PropertyID = at.TagID
		Case EA_MethodTag
    		Dim mt As EA.MethodTag
    		Set mt = m_tv
    		PropertyID = mt.TagID
		Case EA_ConnectorTag	
    		Dim ct As EA.ConnectorTag
    		Set ct = m_tv
    		PropertyID = ct.TagID
		Case EA_RoleTag
    		''' FAKE VALUE: No property ID exist for RoleTag! But we provide a fake ID anyway. 
    		''' In any case the user must check the result when calling GetTaggedValueByID(Id)
    		PropertyID = -1
		Case Else
    		Err.Raise err_TaggedValueType, msg_TaggedValueType 
	End Select		
	End Property
	
	
	''' [FQName]
	''' Fully expanded Stereotype, like so: "Tool::Stereotype::Name"
	Public Property Get FQName() ''': String
	FQName = m_tv.FQName
	End Property
	
	
	''' [M_Value]
	''' Direct value only, no Default values are returned. If you want
	''' Default values if a value isn't set by the user, then call Value()
	''' instead, since it grants that a Default value is returned if it 
	''' by evaluating, in this order: 
	''' 1. Value() || 2. Default() || 3. GlobalDefault()
	Public Property Get M_Value()	''': String
	Dim S
	IncStats()								''' (($stats))
	If m_objecttype = EA_RoleTag Then
		''' Special Case RoleTag
		S = m_rt.Value()
		If TryExtractRoleTagValue(S) Then
			M_Value = S
		Else
			M_Value = ""
		End If
	Else
		''' Normal tags (= all other Tags but RoleTag)
		S = m_tv.Value()
		If  S = "" Then
			M_Value = ""
		ElseIf S = "<memo>" Then
			''' if <memo>, retrive value from .Notes
			M_Value = m_tv.Notes
		Else
			M_Value = S						'// = S already contain the value
		End If
	End If
	End Property
	
	
	''' [M_Default]
	''' Direct value only. The value is retrieved from Stereotypes' "initial 
	''' value" - if any. If no value is found, an attempt to retrived a default 
	''' value from M_GlobalDefault() value instead. But, such semantics is
	''' performed only in the main Value() property.
	Public Property Get M_Default()	''': String
    	Dim S
    	IncStats()								''' (($stats))
    	If TryExtractStereotypeDefault(S) Then
    		M_Default = S
    	Else
    		M_Default = ""
    	End If
	End Property
	
	
	''' [M_GlobalDefault]
	''' Direct value, retrieved from Repository.PropertyTypes, 
	'''	with no extra manipulation of the value is performed.
	Public Property Get M_GlobalDefault()	''': String
    	Dim S
    	Dim sName
    	
    	IncStats()								''' (($stats))
    	
    	''' Get Name for use with  PropertyTypeByName(Name...) below
    	If m_objecttype = EA_RoleTag Then
    		sName = m_rt.Tag
    	Else
    		sName = m_tv.Name
    	End If			
    	
    	If TryGetPropertyTypeDefault(sName, S) Then
    		M_GlobalDefault = S
    	Else
    		M_GlobalDefault = ""
    	End If
	End Property
	
	
	''' [Detail]
	Public Property Get Detail() ''': String
    	Dim p As EA.PropertyType
    	IncStats()									''' (($stats))
    	If PropertyTypeByName(m_tv.Name, p) Then
    		Detail = p.Detail
    	Else
    		Detail = ""
    	End If
	End Property
	
	
	''' [HasNotes]
	''' RoleTags doesn't have any Notes field.
	''' ConnectorEnds, OTOH hand, stores its Notes field in the 
	''' Connector table as t_connector.SourceRoleNote / DestRoleNote
	Public Property Get HasNotes()	''': String
    	IncStats()								''' (($stats))
    	If m_objecttype = EA_RoleTag Then
    		HasNotes = False
    	Else
    		HasNotes = m_tv.Notes <> ""
    	End If
	End Property
	
	
	''' [HasMemo]
	''' RoleTags doesn't have any Memo field.
	Public Property Get HasMemo() ''': Boolean
    	IncStats()								''' (($stats))
    	If m_objecttype = EA_RoleTag Then
    		HasMemo = False
    	Else
    		HasMemo = m_tv.Value = "<memo>"
    	End If
	End Property
	
	
	''' [IsValueDefault]
	''' Determines whether the Value() property is a "native" value or derived from Default()
	Public Property Get IsValueDefault() ''': Boolean
    	IncStats()								''' (($stats))
    	IsValueDefault = (M_Value = "") And (Value <> "")
	End Property
	
	
	''' [ParentType]
	Public Property Get ParentType() ''': String (Kind name)
    	IncStats()								''' (($stats))
    	If m_parent_typename = "" Then
    		Select Case m_objecttype
    			Case EA_TaggedValue		m_parent_typename = Repository.GetElementByID(  	m_tv.ParentID   ).Type
    			Case EA_AttributeTag	m_parent_typename = "Attribute"
    			Case EA_MethodTag		m_parent_typename = "Operation"
    			Case EA_ConnectorTag	m_parent_typename = Repository.GetConnectorByID(	m_tv.ConnectorID).Type
    			Case EA_RoleTag			m_parent_typename = "ConnectorEnd"
    		End Select
    	End If
    	ParentType = m_parent_typename
	End Property
	
	
	''' [ParentID]
	''' Special case for RoleTag / ConnectorEnd
	Public Property Get ParentID() ''': Integer
    	IncStats()								''' (($stats))
    	Select Case m_objecttype
    		Case EA_TaggedValue
        		' Direct use of the internal ref 'm_tv'
        		ParentID = m_tv.ParentID
    		Case EA_AttributeTag	
        		Dim at As EA.AttributeTag
        		Set at = m_tv
        		ParentID = at.AttributeID
    		Case EA_MethodTag
        		Dim mt As EA.MethodTag
        		Set mt = m_tv
        		ParentID = mt.MethodID
    		Case EA_ConnectorTag	
        		''' Return the Connector's ID
        		Dim ct As EA.ConnectorTag
        		Set ct = m_tv
        		ParentID = ct.ConnectorID
    		Case EA_RoleTag
        		''' RoleTagConnector is an expesive call (SQLQuery) 
        		''' but at least it's cached internally
        		ParentID = RoleTagConnector.ConnectorID
    		Case Else
        		Err.Raise err_TaggedValueType, msg_TaggedValueType '// Indicate error
    	End Select			
	End Property
	
	''' [TvObject]
	''' Publishes the currently wrapped TaggedValue. Be aware of that this native 
	''' TV object is NOT "type safe" due to EA tag's inherent un-orthogonality.
	''' Use IsElemenTag, IsConnectorTag, IsRoleTag etc (via this wrapper) to determine 
	'''	the Tag type before using this property.
	Public Property Get TvObject() ''': EA.TaggedValue
		Set TvObject = m_tv
	End Property

		
	''' [ParentObject]
	Public Property Get ParentObject() ''': EA.<Object>
    	IncStats()								''' (($stats))
    	Set ParentObject = M_ParentObject()		''' M_ = No statistics!
	End Property
	
	
	''' [M_ParentObject]
	''' This prop is mainly for internal use. It will NOT succed to call
	''' (un-orthogonal) properties since it returns the actual model entity (thus
	''' for example a RoleTag's Parent (and ConnectionEnd) will not have a property 
	''' Name(), and so an access violation will be rised if calling it.
	'''
	''' ANYWAY, for exactly the above (un-orthogonality) reason, this wrapper also 
	''' provides with type-check functions as enable a convenient means to avoid calling 
	''' unorthogonal properties (See IsRoleTag, IsClassTag, IsPackageTag, IsMethodTag, 
	''' IsAttributeTag and IsConnectionTag, which can be called after the initial Wrap)	
	Public Property Get M_ParentObject() ''': EA.<Object>
    	If m_parentobj Is Nothing Then
    		Select Case m_objecttype
    			Case EA_TaggedValue
    			' Set = use internal ref
    			Set m_parentobj = Repository.GetElementByID(m_tv.ParentID)
    			''' Too expesive to set ParentObject.ObjectType unless required:
    			''' m_parent_objecttype = [skip]
    			Case EA_AttributeTag	
    			Dim at As EA.AttributeTag
    			Set at = m_tv
    			Set m_parentobj = Repository.GetAttributeByID(at.AttributeID)
    			''' While at it, set also:
    			m_parent_objecttype = EA_Attribute
    			Case EA_MethodTag
    			Dim mt As EA.MethodTag
    			Set mt = m_tv
    			Set m_parentobj = Repository.GetMethodByID(mt.MethodID)
    			''' While at it, set also:
    			m_parent_objecttype = EA_Method
    			Case EA_ConnectorTag	
    			''' Return the Connector's name
    			Dim ct As EA.ConnectorTag
    			Set ct = m_tv
    			Set m_parentobj = Repository.GetConnectorByID(ct.ConnectorID)
    			''' While at it, set also:
    			m_parent_objecttype = EA_Connector
    			Case EA_RoleTag
    			''' EA.ConnectorEnd
    			Set m_parentobj = ConnectionEndForRoleTag()
    			''' While at it, set also:
    			m_parent_objecttype = EA_RoleTag
    			Case Else
    			Err.Raise err_TaggedValueType, msg_TaggedValueType '// Indicate error
    		End Select
    	End If
    	Set M_ParentObject = m_parentobj
	End Property
	
	
	''' [ParentName]
	''' This prop. is a "proof of concept" for testing the un-orthogonality 
	''' hidden in this wrapper concept.
	Public Property Get ParentName() ''': String
    	IncStats()									''' (($stats))
    	Select Case m_objecttype
    		Case EA_RoleTag
    		''' Cast neded for access ConnectorEnd's unique properties
    		Dim p_obj As EA.ConnectorEnd
    		Set p_obj = M_ParentObject
    		ParentName = p_obj.Role
    		Case Else
    		ParentName = M_ParentObject.Name
    	End Select
	End Property
	
	
	''' [ParentObjectType]
	Public Property Get ParentObjectType() ''': Integer (ot<ObjectType>)
    	IncStats()									''' (($stats))
    	If m_parent_objecttype > otNone Then		''' otNone = 0
    		M_ParentObjectType = m_parent_objecttype
    	Else
    		''' From the TaggedValue Type we mosty often know which Type the aprent has.
    		''' Check Type of TaggedValue (from Obj.ObjectType which was set already at Wrap)
    		Select Case m_objecttype	
    			Case EA_TaggedValue
    			
    			' Use internal ref
    			Select Case Repository.GetElementByID(m_tv.ParentID).ObjectType
    				Case EA_Package	
    				m_parent_objecttype  = EA_Package
    				Case EA_Element 
    				m_parent_objecttype = EA_Element
    				Case Else 
    				Err.Raise err_ElementType, msg_ElementType '// Indicate error
    			End Select
    			
    			Case EA_AttributeTag
    			m_parent_objecttype = EA_Attribute
    			Case EA_MethodTag
    			m_parent_objecttype = EA_Method
    			Case EA_ConnectorTag
    			m_parent_objecttype = EA_Connector
    			Case EA_RoleTag
    			m_parent_objecttype = EA_ConnectorEnd
    			Case Else 
    			m_parent_objecttype = otNone
    			Err.Raise err_ElementType, msg_ElementType
    		End Select
    	End If
    	M_ParentObjectType = m_parent_objecttype
	End Property
	
	
	''' [IsAttributeTag]
	Public Property Get IsAttributeTag() ''': Boolean
    	IncStats()								''' (($stats))
    	IsAttributeTag = (m_objecttype = EA_AttributeTag)
	End Property
	
	
	''' [IsMethodTag]
	Public Property Get IsMethodTag() ''': Boolean
    	IncStats()								''' (($stats))
    	IsMethodTag = (m_objecttype = EA_MethodTag)
	End Property
	
	
	''' [IsConnectionTag]
	Public Property Get IsConnectionTag() ''': Boolean
    	IncStats()								''' (($stats))
    	IsConnectionTag = (m_objecttype = EA_ConnectorTag)
	End Property
	
	
	''' [IsRoleTag]
	Public Property Get IsRoleTag() ''': Boolean
    	IncStats()								''' (($stats))
    	IsRoleTag = (m_objecttype = EA_RoleTag)
	End Property
	
	
	Public Property Get M_IsRoleTag() ''': Boolean
    	M_IsRoleTag = (m_objecttype = EA_RoleTag)
	End Property
	
	
	''' [IsTaggedValue]
	Public Property Get IsTaggedValue() ''': Boolean
    	IncStats()								''' (($stats))
    	IsTaggedValue = (m_objecttype = EA_TaggedValue)
	End Property
	
	
	''' [IsClassTag]
	''' Element is same as Interface & Class
	Public Property Get IsClassTag() ''': Boolean
    	IncStats()								''' (($stats))
    	IsClassTag = (m_objecttype = EA_TaggedValue)
	End Property
	
	
	''' Element is same as Interface & Class
	Public Property Get IsElementTag() ''': Boolean
    	IncStats()								''' (($stats))
    	IsElementTag = (m_objecttype = EA_TaggedValue)
	End Property
	
	
	''' Element is same as Interface & Class
	''' TODO: Check stereotype to distinguish from regular Class
	Public Property Get IsInterfaceTag() ''': Boolean
    	IncStats()								''' (($stats))
    	IsInterfaceTag = (m_objecttype = EA_TaggedValue)
	End Property
	
	
	''' [IsPackageTag]
	Public Property Get IsPackageTag() ''': Boolean
    	IncStats()								''' (($stats))
    	IsPackageTag = (m_objecttype = EA_TaggedValue)
	End Property
	
	
	''' ----------								
	''' STATS									
	''' ----------					
				
	
	Public Sub StatsStart() ''': Void
		UseStats = True							''' (($stats))
		ResetStats()							''' (($stats))
		m_stats_starttime = Now()				''' (($stats))
	End Sub
	
	
	Public Sub StatsStop() ''': Void
		UseStats = False						''' (($stats))
		m_stats_stoptime = Now()				''' (($stats))
		m_stats_hitcount_acc = m_stats_hitcount_acc + m_stats_hitcount	''' (($stats))
		m_stats_time_acc = m_stats_time_acc + StatsDuration()				''' (($stats))
	End Sub
	
	
	Public Sub StatsPause() ''': Void
		UseStats = False						''' (($stats))
		m_stats_stoptime = Now()				''' (($stats))
		m_stats_hitcount_acc = m_stats_hitcount_acc + m_stats_hitcount	''' (($stats))
		m_stats_time_acc = m_stats_time_acc + StatsDuration()				''' (($stats))
	End Sub
	
	
	Public Sub StatsResume() ''': Void
		UseStats = True									''' (($stats))
		m_stats_hitcount = 0							''' (($stats))
		m_stats_starttime = Now()						''' (($stats))
	End Sub
	
	
	Public Property Get HasStats() ''': Boolean
    	Dim cnt											''' (($stats))
    	cnt = m_stats_hitcount							''' (($stats))
    	cnt = cnt + m_stats_wrapcount					''' (($stats))
    	cnt = cnt + m_stats_hitcount_acc				''' (($stats))
    	HasStats = cnt > 0								''' (($stats))
	End Property
	
	
	Public Property Get StatsDuration() ''': Time
    	If HasStats Then								''' (($stats))
    		StatsDuration = m_stats_stoptime - m_stats_starttime	''' (($stats))
    	Else											''' (($stats))
    		StatsDuration = 0							''' (($stats))
    	End If											''' (($stats))
	End Property
	
	
	Public Property Get StatsDurationAcc() ''': Time
    	If HasStats Then								''' (($stats))
    		StatsDurationAcc = m_stats_time_acc			''' (($stats))
    	Else											''' (($stats))
    		StatsDurationAcc = 0						''' (($stats))
    	End If											''' (($stats))
	End Property
	
	
	Public Property Get StatsCount() ''': Integer
	   StatsCount = m_stats_hitcount					''' (($stats))
	End Property
	
	
	Public Property Get StatsWrapCount() ''': Integer
    	StatsWrapCount = m_stats_wrapcount				''' (($stats))
	End Property
	
	
	Public Property Get StatsCountAcc() ''': Integer
    	StatsCountAcc = m_stats_hitcount_acc			''' (($stats))
	End Property
	
	
	Public Property Get StatsTimePerHits() ''': Time
    	Dim tmp											''' (($stats))
    	If HasStats And (m_stats_hitcount > 0) Then     ''' (($stats))
    		tmp = (m_stats_stoptime - m_stats_starttime) / m_stats_hitcount	''' (($stats))
    	Else											''' (($stats))
    		tmp = 0										''' (($stats))
    	End If											''' (($stats))
    	StatsTimePerHits = tmp							''' (($stats))
	End Property
	
	
	Public Property Get StatsTimePerHitsAcc() ''': Integer
    	Dim tmp											''' (($stats))
    	If HasStats And (m_stats_hitcount_acc> 0) Then				''' (($stats))
    		tmp = m_stats_time_acc / m_stats_hitcount_acc	''' (($stats))
    	Else											''' (($stats))
    		tmp = 0										''' (($stats))
    	End If											''' (($stats))
    	StatsTimePerHitsAcc = tmp						''' (($stats))
	End Property
	
	
	Public Property Get StatsHitsPerSecond() ''': Integer
    	Dim tmp											''' (($stats))
    	tmp = 0											''' (($stats))
    	If HasStats Then								''' (($stats))
    		tmp = StatsDuration							''' (($stats))
    		If tmp > 0 Then								''' (($stats))
    			tmp = StatsCount / Second(tmp)			''' (($stats))
    		Else										''' (($stats))
    			tmp = 0									''' (($stats))
    		End If										''' (($stats))
    	Else											''' (($stats))
    		tmp = 0										''' (($stats))
    	End If											''' (($stats))
    	StatsHitsPerSecond = tmp						''' (($stats))
	End Property
	
	
	Public Property Get StatsHitsPerSecondAcc() ''': Integer
    	Dim acc_cnt										''' (($stats))
    	Dim acc_dur										''' (($stats))
    	Dim res											''' (($stats))
    	
    	res = 0.0										''' (($stats))
    	acc_dur = StatsDurationAcc()					''' (($stats))
    	If HasStats And (acc_dur > 0) Then				''' (($stats))
    		acc_cnt = StatsCountAcc() 					''' (($stats))
    		On Error Resume Next
    		
    		If Second(acc_dur) > 0 Then					''' (($stats))
    			res = acc_cnt / Second(acc_dur)			''' (($stats))
    		End If										''' (($stats))
    		If Err Then									''' (($stats))
    			'''
    		End If										''' (($stats))
    	End If											''' (($stats))
    	StatsHitsPerSecondAcc = res						''' (($stats))
	End Property
	
	
	''' Private -------------
	
	Private Property Let UseStats(aBool) ''': Void	''' (($stats))
    	m_usestats = aBool								''' (($stats))
	End Property										''' (($stats))
	
	
	''' (($stats))
	Private Sub ResetStats() ''': Void					''' (($stats))
		m_stats_hitcount 		= 0						''' (($stats))
		m_stats_hitcount_acc 	= 0						''' (($stats))
		m_stats_starttime		= 0.0					''' (($stats))
		m_stats_stoptime 		= 0.0					''' (($stats))
		m_stats_time_acc 		= 0.0					''' (($stats))
		m_stats_wrapcount 		= 0						''' (($stats))
	End Sub												''' (($stats))
	
	
	''' (($stats))
	Private Sub IncStats() ''': Void					''' (($stats))
		If m_usestats Then 								''' (($stats))
			m_stats_hitcount = m_stats_hitcount + 1	''' (($stats))
		End If											''' (($stats))
	End Sub												''' (($stats))
	
	
	''' --------------------------
	''' PUBLIC FUNCTIONS
	''' --------------------------
	
	''' [Update]
	''' Updates internal Dictionaries by emptying 
	''' and then re-importing data. 
	''' Called also from Class_Initialize()
	Public Sub Update() ''': Void
		ResetData()
		LoadPropertyData()			
	End Sub
	
	
	''' [ResetData]
	''' Updates internal Dictionaries by emptying them
	''' Called from Update() and Class_Initialize()
	Private Sub ResetData() ''': Void
		m_parent_typename 	= ""
		m_objecttype		= otNone
		m_parent_objecttype	= otNone
		m_startpos			= 0
		m_endpos			= 0
		
		Set m_tv = Nothing
		Set m_rt = Nothing
		Set m_parentobj = Nothing
		Set m_roletag_connector = Nothing
		
		Set m_PropertyTypesDictionary = Nothing
		Set m_PropertyTypesDefaultDictionary = Nothing
		Set m_PropertyTypesRawDataDictionary = Nothing
		
		m_usestats			= False						''' (($stats))
		ResetStats()									''' (($stats))			
	End Sub
	
	
	''' [LoadPropertyData]
	''' Updates internal Dictionaries by re-importing data. 
	''' Called also from Update()  and Class_Initialize()
	Private Sub LoadPropertyData()
		RegisterPropertyTypes()
		RegisterPropertyTypesRawData()
		RegisterPropertyTypesDefaults()		
	End Sub
	
	
	''' --------------------------
	''' INTERNAL PROPERTIES
	''' --------------------------
	
	''' [ConnectionEndForRoleTag]
	''' Returns the ConnectorEnd Object
	''' Accessing a ConnectorEnd/RoleObj from TaggedValues must be done via 
	''' its Connector (since ConnectorEnds are stored in the same table, 
	''' the 't_connector'
	''' Determine which end of the (parent) Connector to read from
	''' Notice that the property RoleTagConnector is cached.
	Private Property Get ConnectionEndForRoleTag()	
    	If m_rt Is Nothing Then
    		Err.Raise err_TaggedValueType, msg_TaggedValueType '// Indicate error
    	ElseIf IsClient Then  	''' = EA_ASSOCIATION_SOURCE
    		Set ConnectionEndForRoleTag = RoleTagConnector.ClientEnd
    	Else				''' = EA_ASSOCIATION_TARGET
    		Set ConnectionEndForRoleTag = RoleTagConnector.SupplierEnd
    	End If
	End Property
	
	
	''' [ConnectorForRoleTag]
	Private Property Get ConnectorForRoleTag()
    	Dim guid
    	If QueryRoleTagForElementID(guid) Then 
    		Set ConnectorForRoleTag = Repository.GetConnectorByGuid(guid)
    	Else
    		Set ConnectorForRoleTag = Nothing
    	End If
	End Property
	
	
	''' Applies only when RoleTag is visitor. 
	''' Raises an error if TaggeValue Visitor is not of the type EA.RoleTag
	''' Only for internal use as to support RoleTags with optimized acceess
	''' to it's owning Connector (because many RoleTag properties are stored
	''' in the owning Connector's t_cmnnector table)
	Private Property Get RoleTagConnector()
    	If m_rt Is Nothing Then
    		Err.Raise err_TaggedValueType, msg_TaggedValueType
    	ElseIf m_roletag_connector Is Nothing Then
    		Set m_roletag_connector = ConnectorForRoleTag()
    	End If
    	Set RoleTagConnector = m_roletag_connector
	End Property
	
	
	''' [PropertyTypesDictionary]
	''' Stores the actual PropertyTypes from Repository.PropertyTypes
	''' for fast access. TaggedValue Name is Key, the TV Object is data
	Private Property Get PropertyTypesDictionary
    	If m_PropertyTypesDictionary Is Nothing Then _
    	Set m_PropertyTypesDictionary = CreateObject("Scripting.Dictionary")
    	Set PropertyTypesDictionary = m_PropertyTypesDictionary
	End Property
	
	
	''' [PropertyTypesDefaultDictionary]
	''' Stores the cached DEFAULT VALUE of each PropertyTypes from the Repository.PropertyTypes
	''' for fast access. TaggedValue Name is Key, the Default Value is data
	Private Property Get PropertyTypesDefaultDictionary()  ''': Dictionary
    	If m_PropertyTypesDefaultDictionary Is Nothing Then _
    	   Set m_PropertyTypesDefaultDictionary = CreateObject("Scripting.Dictionary")
    	Set PropertyTypesDefaultDictionary = m_PropertyTypesDefaultDictionary
	End Property
	
	
	''' [PropertyTypesRawDataDictionary]
	''' Stores the cached RAW TEXT from each PropertyTypes from the Repository.PropertyTypes
	''' for fast access. TaggedValue Name is Key, the raw text is data
	Private Property Get PropertyTypesRawDataDictionary() ''': Dictionary
    	If m_PropertyTypesRawDataDictionary Is Nothing Then
    		Set m_PropertyTypesRawDataDictionary = CreateObject("Scripting.Dictionary")
    	End If
    	Set PropertyTypesRawDataDictionary = m_PropertyTypesRawDataDictionary
	End Property
	
	
	''' [IsSource]
	Private Property Get IsSource() ''': Boolean
    	IsSource = m_rt.BaseClass = EA_ASSOCIATION_SOURCE
	End Property
	
	
	''' [IsClient]
	Private Property Get IsClient() ''': Boolean
	   IsClient = m_rt.BaseClass = EA_ASSOCIATION_SOURCE
	End Property
	
	
	''' [IsTarget]
	Private Property Get IsTarget() ''': Boolean
	   IsTarget = m_rt.BaseClass = EA_ASSOCIATION_TARGET
	End Property
	
	
	''' [IsSupplier]
	Private Property Get IsSupplier() ''': Boolean
	   IsSupplier = m_rt.BaseClass = EA_ASSOCIATION_TARGET
	End Property
	
	
	''' [TryExtractStereotypeDefault]
	''' This level of default retrieves its source data
	''' from the Stereotype Initial value, which will have 
	''' the following format: "...Default:<Value>"
	Private Function TryExtractStereotypeDefault(ByRef s) ''': Boolean		
		''' Special case for RoleTags
		If m_objecttype = EA_RoleTag Then		
			If TryExtractRoleTagStereotypeDefault(s) Then
				TryExtractStereotypeDefault = s <> ""
			Else
				s = ""
				TryExtractStereotypeDefault = False
			End If			
		Else
			Dim tmp
			m_startpos = 0
			tmp = m_tv.Notes()
			
			m_startpos = InStr(1, tmp, "Default:", 1)
			''' If Contains :
			If m_startpos > 0 Then
				m_startpos = m_startpos + 8  ''' = Len("Default:")
				''' Use case-IN-sensitive search and strip out 
				''' the value part to the right of the text "Default:"
				s = Trim(Mid(tmp, m_startpos ))
				TryExtractStereotypeDefault = s <> ""
			Else
				s = ""
				TryExtractStereotypeDefault = False
			End If
		End If
	End Function
	
	
	'''	Getters for RoleTag
	
	
	''' [TryExtractRoleTagValue]
	''' Returns pnly the Value(), if any. Since this is only a "helper" function 
	''' thus it must not control the semantics of values. 
	''' However, it actually returns the string content "$ea_notes=", although the 
	'''	function result returns "False" as to leave to the caller to determine 
	''' whether to display that (control) string or not.
	Private Function TryExtractRoleTagValue(ByRef S) ''': Boolean
		''' ------------------------------------
		''' FUTURE FUNCTIONALITY
		''' If m_rt.HasAttributes() then 
		''' 	m_rt.GetAttribute("$ea_notes")
		''' ------------------------------------
		''' Example-string to examine; "Value$ea_notes=Default: DefaultValue"
		m_startpos = InStr(1, S, "$ea_notes=", 1)
		If m_startpos > 1 Then
			''' Contains 'Value' (which has precedence over 'DefaultValue')
			S = Trim( Mid(S, 1, m_startpos-1) )
			TryExtractRoleTagValue = True
		Else
			TryExtractRoleTagValue = False
		End If
	End Function
	
	
	''' [TryExtractRoleTagStereotypeDefault]
	''' Returns ONLY Default value (from Stereotype) and disregards any Value() or 
	''' GlobalDefault() value. 
	''' Strict "m_startpos = 1" logic will NOT return the default value even if it
	''' exists, if a value is preceeding it. Think about that.
	''' 
	''' In this case we only want a default value from here if - and only if - no 
	''' value is present in front of the "$ea_notes=" control string. The reason for 
	''' this is that the user has selected a another value than the 
	''' default value for this TV (store BEFORE the $ea_notes tag), and such explicit 
	''' user choices must never be overrided.
	''' 
	''' For "direct access" of the property content (without any semantics applied), 
	''' use the M_Value() or M_Default() instead.
	Private Function TryExtractRoleTagStereotypeDefault(ByRef S) ''': String, Boolean
		''' Example content :
		''' S = Value$ea_notes=Default: DefaultValue   
		m_startpos = InStr(1, S, "$ea_notes=Default:", 1)
		If m_startpos = 1 Then
			''' Try extracting to the right of "$ea_notes=Default:"
			m_startpos = m_startpos + 18 	''' 18 = Len("$ea_notes=Default:")
			S = Trim( Mid(S, m_startpos) )
			TryExtractRoleTagStereotypeDefault = True
		Else
			S = ""
			TryExtractRoleTagStereotypeDefault = False
		End If		
	End Function
	
	
	''' [PropertyTypeByName]
	Private Function PropertyTypeByName(aNameAsKey, ByRef OutProp) ''': PropertyType, Boolean
		''' OutResult type: EA.PropertyType
		If m_PropertyTypesDictionary.Exists(aNameAsKey) Then
			Set OutProp = m_PropertyTypesDictionary(aNameAsKey)
		Else
			Set OutProp = Nothing
		End If
		PropertyTypeByName = Not OutProp Is Nothing
	End Function
	
	
	''' [TryGetPropertyTypeDefault]
	Private Function TryGetPropertyTypeDefault(aNameAsKey, ByRef OutResult) ''': String, Boolean
		If m_PropertyTypesDefaultDictionary.Exists(aNameAsKey) Then
			OutResult = m_PropertyTypesDefaultDictionary(aNameAsKey)
			TryGetPropertyTypeDefault = OutResult <> ""
		Else
			OutResult = ""
			TryGetPropertyTypeDefault = False
		End If			
	End Function
	
	
	''' [RegisterPropertyTypes]
	''' Stores the very PropertyType object with the name as Key.
	Private Sub RegisterPropertyTypes() ''': Void
		Dim pt As EA.PropertyType
		
		For Each pt In Repository.PropertyTypes
			PropertyTypesDictionary.Add pt.Tag, pt
		Next
	End Sub
	
	
	''' [RegisterPropertyTypesDefaults]
	''' Stores the refined, extracted default value, as defined in 
	''' the PropertyType text "blob", into a Dictionary. The  
	''' name is used as the Key.
	Private Sub RegisterPropertyTypesDefaults() ''': Void
		Dim sRawData
		Dim sDefault
		Dim sPropNameKey
		
		For Each sPropNameKey In PropertyTypesRawDataDictionary.Keys
			''' Prepare / refine raw data before inserting it into the Dictionary
			sRawData = m_PropertyTypesRawDataDictionary(sPropNameKey)
			sDefault = ExtractPropertyFromRawStr(sRawData, "Default")
			
			'''Add to dictionary
			PropertyTypesDefaultDictionary.Add sPropNameKey, sDefault
		Next
	End Sub
	
	
	''' [RegisterPropertyTypesRawData]
	Private Sub RegisterPropertyTypesRawData() ''': Void
		Dim p As EA.PropertyType	
		Dim sTemp
		For Each p In Repository.PropertyTypes
			sTemp = p.Detail()
			''' Prepare the raw str for faster extraction (Lazy Eval)
			FormatPropertyTypesText sTemp 
			''' Store so it can later be retrieved ByName (= p.Tag)
			PropertyTypesRawDataDictionary.Add p.Tag, sTemp
		Next
	End Sub
	
	
	''' HELPER FUNCTIONS
	
	
	''' [GetValueByXmlTagName]
	''' Used for extracting single values from Repository.SQLQuery results (xml format).
	''' Omit the <> tags in aTagName.
	Private Function GetValueByXmlTagName(ByRef aStr, ByRef aTagName, ByRef OutResult) ''': String, Boolean
		Dim sTag
		
		m_startpos = 0
		m_endpos = 0
		sTag = "<" & aTagName & ">"
		''' Get first tag pos
		m_startpos = InStr(1, aStr, sTag, 1) + Len(aTagName)+2
		If m_startpos > 0 Then
			''' End tag pos
			m_endpos = InStr(1, aStr, "</" & aTagName & ">", 1)
			'' The value
			OutResult = Mid(aStr, m_startpos, m_endpos - m_startpos)
			GetValueByXmlTagName = True
		Else
			OutResult = ""
			GetValueByXmlTagName = False
		End If
	End Function
	
	
	''' [QueryRoleTagForElementID]
	''' Retrieves ElementID from EA.TagRole stored in "t_taggedvalue".  This 
	''' value is not exposed by the EA.RoleTag although there's a method for it.
	''' Be aware of that the property name (PropertyGUID) is not the same as 
	''' the table name (PropertyID).
	''' The resulting GUID to be used for fetching the parent Connector 
	''' from the table "t_connector". Code example:
	''' -----------------------------------------------------------------------
	''' If QueryRoleTagForElementID(m_rt, guid) Then _
	''' 	Set conn = Repository.GetConnectorByGuid(guid)
	''' -----------------------------------------------------------------------
	Private Function QueryRoleTagForElementID(ByRef OutGUID) ''': Boolean
		Dim result
		''' SQL - The query returns xml, in this format:
		''' -------------------------------------------------------------------
		'''		<EADATA version="1.0" exporter="Enterprise Architect">
		'''			<Dataset_0><Data><Row><ElementID>{D5C40150-0CE8-4c24-A635-C508623F9D45}</ElementID></Row></Data></Dataset_0></EADATA>	
		''' -------------------------------------------------------------------
		result = Repository.SQLQuery( _
    		"SELECT t_taggedvalue.ElementID " & _
    		"FROM t_taggedvalue " & _
    		"WHERE (PropertyID='" & m_rt.PropertyGUID & "');" _
    		)
		If GetValueByXmlTagName(result, "ElementID", result) And (result<>"") Then
			OutGUID = result
			QueryRoleTagForElementID = True
		Else
			OutGUID = ""
			QueryRoleTagForElementID = False
		End If
	End Function
	
	
	''' ----------------------------------
	''' EXTRACT (ANY) PROPERTY FROM STRING
	''' ----------------------------------
	''' [ExtractPropertyFromRawStr]
	''' EXTRACT VALUE FROM STRING
	''' * Result: Extracts named values from a "lump" field of multiple named 
	'''   property values. 
	''' * Entry : aSubjectStr MUST be treated with "FormatPropertyTypesText" before 
	'''	  calling this. Using Case IN-sensitive search for keywords.
	''' Example value as defined for an individual PropertyType (TaggedValue):
	'''  "Persistence : Type=Enum;
	'''   Values=Persistent;Transient;
	'''   Default=Persistent;
	'''   BaseStereotype=class;attribute;...etc;"
	Private Function ExtractPropertyFromRawStr(ByRef aSubjectStr, ByVal aFieldName) ''': Boolean
		Const vbCaseInSensitive = 1
		Dim pos_crlf
		Dim copy_len
		Dim eol_char
		Dim delimiter
		Dim result
		
		If aFieldName="" Then
			MsgBox "Error: Search string is empty!"
			ExtractPropertyFromRawStr = ""
			Exit Function
		End If
		
		aSubjectStr = aSubjectStr
		If InStr(1, aSubjectStr, ";", vbCaseInSensitive) > 0 Then
			MsgBox "Error: Subject string was not formatted using the method 'FormatPropertyTypesText'"
			ExtractPropertyFromRawStr = ""
			Exit Function
		End If
		
		result = ""
		
		''' Ensure traling "=" char to search for
		If Right(aFieldName, 1) <> "=" Then _
		aFieldName = aFieldName & "="
		
		m_startpos = 0
		m_startpos = InStr(1, aSubjectStr, aFieldName, 1)	''' 1 = Case Insensitive
		
		''' Check if value is "terminated" (=has a trailing ";" after the match)
		If m_startpos > 0 Then
			m_startpos = m_startpos + Len(aFieldName)
			
			''' Find the end of line
			eol_char = Chr(10)
			m_endpos = InStr(m_startpos, aSubjectStr, eol_char, 1)
			
			''' Check if a CRLF is located *before* the ";"
			pos_crlf = InStr(m_startpos, aSubjectStr, Chr(10), 1)
			
			''' If no CRLF exist, or is located *before* the ";" then that 
			''' (probably) means that a semicolon has been used as a delimiter 
			''' between values, and not only as a EOL char. Therefore, use LF 
			''' as the terminating char instead, and update the delimiter, as 
			''' well as the m_endpos (of line) accordingly.
			
			If (pos_crlf = 0) Or (pos_crlf > m_endpos) Then
				''' Update EOL pos
				m_endpos = pos_crlf
			End If
			
			''' Extract the value part
			If m_endpos > m_startpos Then 
				''' .......................................................
				''' Advance the start pos with the length of the search str
				''' Example: "... Default=<some value>; ... "
				'''               |    -->|
				''' .......................................................
				copy_len = m_endpos - m_startpos
				
				''' Extract
				result = Trim(Mid(aSubjectStr, m_startpos, copy_len))
			End If
		Else
			result = ""
		End If
		ExtractPropertyFromRawStr = result
	End Function
	
	
	''' [ContainsStr]
	''' Case INsensitive search
	Private Function ContainsStr(aStr, aChar)
		ContainsStr = InStr(1, aStr, aChar, 1) > 0	''' vbTextCompare
	End Function
	
	

    ''' [FormatPropertyTypesText]
    ''' Removes any delimiters from line ends, keeping 
    ''' only the LF. Thereafter all delimiters, like ";", 
    ''' are replaces with commas.
    ''' Ensures that the last line is treated like the other lines,
    ''' also meaning that Chr(10) (single LF) can be used as the 
    ''' terminating character.
    Private Sub FormatPropertyTypesText(ByRef aSubjectStr)
    	aSubjectStr = aSubjectStr & Chr(10)
    	
    	''' Consistent formatting for the rest (, as delimiter, and no ";" 
    	''' at the ned of lines) : 
    	
    	aSubjectStr = Replace(aSubjectStr, Chr(13), Chr(10), 1, -1, 1)
    	aSubjectStr = Replace(aSubjectStr, ";" & Chr(10), Chr(10), 1, -1, 1)
    	aSubjectStr = Replace(aSubjectStr, "," & Chr(10), Chr(10), 1, -1, 1)
    	aSubjectStr = Replace(aSubjectStr, ";", ",", 1, -1, 1)
    End Sub
	
	''' TRILTaggedValueApi
End Class

Module_Initialize