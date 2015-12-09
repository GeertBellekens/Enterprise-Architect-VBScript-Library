
#**TaggedValue Helper for EA**

<dl>
<table>
<tr><td>Author:</td><td>Rolf Lampa, rolf.lampa@rilnet.com</td></tr>
<tr><td>Copyright:</td><td>(C) Rolf Lampa, 2015. This script is free to use for commercial projects provided that proper attribution is given to the author, and providing this copyright info in a visible place in your code and product documentation, including a link to this page.</td></tr>
</table>
</dl>

*For some information about the background for this TaggedValue Helper, see a discussion  thread at EA's forum ni [this link](http://www.sparxsystems.com/cgi-bin/yabb/YaBB.cgi?num=1448991338/15#15 "Chart of Different Property Names, etc")*

####TOC
* [Providing Consistent Property Names](README.md#providing-consistent-property-names)<br>
* [Better Properties](README.md#better-properties)<br>
* [More Properties](README.md#more-properties)<br>
* [EXAMPLE OF USAGE](README.md#example-of-usage)<br>
* [Try or just nail it?](README.md#try-or-just-nail-it)<br>
* [DETAILED DESCRIPTION](README.md#detailed-description)<br>
* [Fake "Polymorphism"](README.md#fake-polymorphism)<br>
* [Statistics](README.md#statistics)<br>
* [Todo](README.md#todo)<br>
* [CLASS MEMBERS](README.md#class-members)<br>
* [Donations](README.md#donations)<br>

####**Providing Consistent Property Names**
<img src="http://wiki.rilpartner.se/w/images/wiki.rilpartner.se/2/2d/EA-TaggedValue-System.jpg" 
alt="TaggedValue system" align="right" width="540" border="10"/>
This TaggedValue Helper wrapper for Enterprise Architect, written in VBScript, intends to provide advanced users of EA with simpler access to TaggedValue properties with a set of consistent property names, an orthogonality which, as of this writing, is lacking in EA regarding the TaggedValues system (see tables below about the inconsistent naming of the properties in the EA API). On this page the word `Tag`, or the acronym `TV`, may occasionally be used instead of `TaggedValue`.

####**Better Properties**
One of the most important features of the wrapper which has been added, is the much smarter Value() property, which automagically delivers any Default() values, if any such default value was defined in your own **`<<Stereotypes>>`**'  "initial value" field, or as a last alternative if no other value was defined, in the "global" TaggedValue definitions stored in **`PropertyTypes`** (called Project | Settings | "**UML Types**" in the UI).

The way one retrieves these values from the inner workings of EA are *also different* for some of the Tag types, and in some cases these values requires complex programming in order to be accessed, but again, this can't be done in a consistent manner using the EA API.  But the good news is that this helper wrapper does all this for you while hiding all the complexity. And it does more than so.

####**More Properties**
Although differencies in the API exists between Tag types such as `EA.PackageTag`, `EA.ElementTag`, `EA.AttributeTag`, `EA.MethodTag` and `EA.ConnectorTag`, and most different of all, the EA.RoleTag, the differences disappears altogether when using this wrapper. And furthermore, the wrapper even provides with a set of more useful properties than the API, properties that can significantly simplify access to model info which would require sometimes quite complex coding, such as retrieval of parent info (for the `TaggedValues`), following *pathways which also is different* for different Tag types and especially for `RoleTags`, which can prove to be even difficult.

A more detailed description can be found far below on this page, but for coders, let's get right at it with an example of how the wrapper is used, making your life easier dealing with TaggedValues with VBscript:

#####**EXAMPLE OF USAGE:**#####
```vbs
Dim tv As EA.TaggedValue
For Each tv in Pkg.Element.TaggedValues
    Session.Output TagAPI.Wrap(tv)           // Wrap!
    Session.Output TagAPI.Wrap(tv).Value()   // or Wrap and use directly
    Session.Output TagAPI.Name()             // Now use the helper Obj directly
    Session.Output m_tagapi.Notes()          // Or use the module variable
    Session.Output m_tagapi.FQName()
    Session.Output m_tagapi.PropertyGUID()
    Session.Output m_tagapi.ParentName()
    Session.Output m_tagapi.ParentID()
    Session.Output m_tagapi.ParentGUID()
    ''' Etc
Next
```
**One can also wrap the parent class** of the Tag and directly pick the desired Tag "ByName". The wrapper then looks up the TaggedValue in the parent's `TaggedValues` collection, wraps it, and immediately exposes its properties, like so:
```vbs
For Each elem in Pkg.Elements
    If TagApi.TryWrapByName("VBA.FileName", elem)  then
    Session.Output m_tagapi.Name()
    Session.Output m_tagapi.Notes()
    Session.Output m_tagapi.PropertyGUID()
    End If
Next
```

Or, if assuming that the TV really exists, use the more "direct" version of the function, which returns the wrapper object, which then also directly provides the properties of the desired TaggedValue, in this example, the `.Notes()` property:

```vbs
S = TagApi.WrapByName("VBA.FileName", elem).Notes()
```
####**Try or just nail it?**
**Notice how** the `TryWrapByName()` method above is a Function returning `Boolean` in order to assist the programmer in determining whether any value or Tag with that name was found, whereas the WrapByName() returns the wrapper itself directly (no questions answered). In the latter case it will of course end up in an access violation if the named Tag turns out to not exist, but sometimes you just know what you have. But by reason of the uncertain cases the TryWrapByName function may be a better choice, at least if you are uncertain of whether the named Tag actually exist.

###**DETAILED DESCRIPTION**

**Missing, Renamed and Enhanced** - Some properties in EA's Api are "missing", but are published in the wrapper. This is useful when traversing the model structure and generic object names makes such loops easier. Some other properties have different names in EA's API, which the wrapper of course aligns using consistsnt naming for all Tag types. And the most important property of the all, the `Value()` property, derives its value (if it's own direct value is not specified) from Default values specified in other places, if any default values are defined at all. 
 
####**"Fake Polymorphism"**
<img src="http://wiki.rilpartner.se/w/images/wiki.rilpartner.se/0/03/TaggedValue_Default_from_Stereotype.jpg" 
alt="PropertyTypes / UML Types" align="right" width="480" border="10"/>
**The derivation follows this order** : If no *direct* value is specified a a Value() then it attempts to derive a default value from #1  `Stereotype's` "initial value" (if defined) for a given TaggedValue. 

But if no default value is defined in the Stereotype then it #2 attempts to derive a default value from `Repository.PropertyTypes` instead, see fig.3 (In the text below this default value is sometimes called "global" Default value). <img src="http://wiki.rilpartner.se/w/images/wiki.rilpartner.se/3/3d/Value_and_Default_Value_in_UML_Types_-_VBA.VBAName.jpg" alt="PropertyTypes / UML Types" align="right" width="480" border="10"/> 

Only if no default values are defined and no direct value is specified (by the user), only then the Value() property gives up its derivation attempts and returns an empty string instead.
 
**An important extra feature** is that the wrapper class also provides easy access to the TaggedValue's parent objects, such as Classes, Attributes and, most difficult to access, ConnectionEnd objects for RoleTags. Property names are "aligned" to be close to similar to the property names of the EA.TaggedValue type, and the user of the class will never have to care about whether the `TaggedValue` (owned by classes, packages and interfaces) is an `AttributeTag`, `MethodTag`, `ConnectorTag` or `RoleTag` since all properties are the same (orthogonal).

###**STATISTICS**

Support for collecting some basic runtime statistics about the number of times TaggedValues has been accessed and time spent evaluating them, has also been implemented in the wrapper. This functionality can be "disabled" from the code base altogether by using the following [Regular Expression](http://www.regular-expressions.info) (tested with **EditPad Lite**, a free version can be downloaded from [here](http://www.editpadlite.com/download.html "EditPad's Download page") ). Expressions to be used are the following:

######**DISABLE** the Stats code in the source code (by commenting):
	Regex Search:		^(?!'//)(.*?\(\(\$stats\)\).*?$)
	Regex Replace:		'//\1
######**ENABLE** the `Stats` code rows in the source (removes commenting):
	Regex Search:		^(?='//)'//(.*?\(\(\$stats\)\).*?$)
	Regex Replace:		\1

#####**TODO**
    Property `IsInterfaceTag()` - Needs checking of the Stereotype in order to be 
                                  distinguished from a regular Class.

----
##**CLASS MEMBERS**
First the most frequently used Properties & Functions, and below that a full list of public members:

```vbs
Class TaggedValue
Public Function Wrap(ByRef aTaggedValue) ''': EA.TaggedValue (the wrapped Tag);
''' Passing Tag's parent object as "aObj" causes a lookup of the desired 
''' TaggedValue (returns True if the TV exists). The TV's properties are available.
Public Function TryWrapByName(aName, ByRef aObj) ''': Boolean; 	Passing the Tag's 
''' parent object as "aObj" gives direct access to properties (assuming that the 
''' TV actually exists). 
Public Function WrapByName(aName, ByRef aObj) ''': TaggedValue (this Wrapper class)
Public Property Get Value() ''': String
Public Function TryValue(ByRef S) ''': Boolean
Public Property Get Name() ''': String
Public Property Get Notes() ''': String
Public Sub Update() ''': Void			''' All PropertyTypes reloaded from
''' the EA Repository in a total re-initialization
```
Other useful and orthogonal properties:

```vbs
''' In case a value isn't actually provided by the underlaying 
''' object, these properties at least provides with a fake value  
''' allowing for "type safe" traversing of EA models.

''' PUBLIC

Public Property Get Detail() ''': String
Public Property Get FQName() ''': String
Public Property Get HasMemo() ''': Boolean
Public Property Get HasNotes()	''': String
Public Property Get HasStats() ''': Boolean
Public Property Get IsAttributeTag() ''': Boolean
Public Property Get IsClassTag() ''': Boolean
Public Property Get IsConnectionTag() ''': Boolean
Public Property Get IsElementTag() ''': Boolean
Public Property Get IsInterfaceTag() ''': Boolean
Public Property Get IsMethodTag() ''': Boolean
Public Property Get IsPackageTag() ''': Boolean
Public Property Get IsRoleTag() ''': Boolean
Public Property Get IsTaggedValue() ''': Boolean
Public Property Get IsValueDefault() ''': Boolean
Public Property Get M_Default()	''': String
Public Property Get M_GlobalDefault()	''': String
Public Property Get M_IsRoleTag() ''': Boolean
Public Property Get M_ParentObject() ''': EA.<Object>
Public Property Get M_Value() ''': String
Public Property Get TvObject() ''': EA.TaggedValue		''' Useful with WrapByName
Public Property Get ParentID() ''': Integer
Public Property Get ParentName() ''': String
Public Property Get ParentObject() ''': EA.<Object>
Public Property Get ParentObjectType() ''': Integer (ot<ObjectType>)
Public Property Get ParentType() ''': String (Kind name)
Public Property Get PropertyGUID() ''': String
Public Property Get PropertyID() ''': String

''' STATS

Public Property Get StatsCount() ''': Integer
Public Property Get StatsCountAcc() ''': Integer
Public Property Get StatsDuration() ''': Time
Public Property Get StatsDurationAcc() ''': Time
Public Property Get StatsHitsPerSecond() ''': Integer
Public Property Get StatsHitsPerSecondAcc() ''': Integer
Public Property Get StatsTimePerHits() ''': Time
Public Property Get StatsTimePerHitsAcc() ''': Integer
Public Property Get StatsWrapCount() ''': Integer

''' PRIVATE

Private Function ContainsStr(aStr, aChar)
Private Function ExtractPropertyFromRawStr(ByRef aSubjectStr, ByVal aFieldName) ''': Boolean
Private Function GetValueByXmlTagName(ByRef aStr, ByRef aTagName, ByRef OutResult) ''': String, Boolean
Private function PropertyTypeByName(aNameAsKey, ByRef OutProp) ''': PropertyType, Boolean
Private Function QueryRoleTagForElementID(ByRef OutGUID) ''': Boolean
Private Function TryExtractRoleTagStereotypeDefault(ByRef S) ''': String, Boolean
Private Function TryExtractRoleTagValue(ByRef S) ''': Boolean
Private Function TryExtractStereotypeDefault(ByRef s) ''': Boolean
Private function TryGetPropertyTypeDefault(aNameAsKey, ByRef OutResult) ''': String, Boolean
Private Property Get ConnectionEndForRoleTag()
Private Property Get ConnectorForRoleTag()
Private Property Get IsClient() ''': Boolean
Private Property Get IsSource() ''': Boolean
Private Property Get IsSupplier() ''': Boolean
Private Property Get IsTarget() ''': Boolean
Private Property Get PropertyTypesDefaultDictionary()  ''': Dictionary
Private Property Get PropertyTypesDictionary
Private Property Get PropertyTypesRawDataDictionary() ''': Dictionary
Private Property Get RoleTagConnector()
Private Property Let UseStats(aBool) ''': Void	''' (($stats))

Private Sub Class_Initialize() ''': Void
Private Sub Class_Terminate() ''': Void

Private Sub FormatPropertyTypesText(ByRef aSubjectStr)
Private Sub IncStats() ''': Void				''' (($stats))
Private Sub LoadPropertyData()
Private Sub RegisterPropertyTypes() ''': Void
Private Sub RegisterPropertyTypesDefaults() ''': Void
Private Sub RegisterPropertyTypesRawData() ''': Void
Private Sub ResetData() ''': Void
Private Sub ResetStats() ''': Void				''' (($stats))
End Class
```

###**Donations**
Although we love to provide useful things for free saving you lots of time and hassle, we also spend lots of time making the life easier for EA developers. If you find the script being useful you may consider making a donation. All amounts amounts. For Paypal donations, use the following 
[Paypal Link](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=2VFSWN93XEPZ2 "Paypal's Secure Pages")

<dl>
<form action="https://www.paypal.com/cgi-bin/webscr" method="post" target="_top">
<input type="hidden" name="cmd" value="_s-xclick">
<input type="hidden" name="hosted_button_id" value="KJCD6N8M8MRWQ">
<input type="image" src="https://www.paypalobjects.com/en_US/SE/i/btn/btn_donateCC_LG.gif" border="0" name="submit" alt="PayPal - The safer, easier way to pay online!">
<img alt="" border="0" src="https://www.paypalobjects.com/en_US/i/scr/pixel.gif" width="1" height="1">
</form>
</dl>
