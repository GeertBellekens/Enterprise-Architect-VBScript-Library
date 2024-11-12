'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit

!INC Local Scripts.EAConstants-VBScript

' EA-Matic
' Script Name: BIGroupAttributes
' Author: Matthias Van der Elst
' Purpose: Set SourceAttribute & Stereotype for BI-Group attributes
' Date: 02/05/2017
'

dim sourceAttribute 'The source attribute

function EA_OnContextItemChanged(GUID, ot)
  if ot = otAttribute then
  'get the attribute
  set sourceAttribute = Repository.GetAttributeByGuid(GUID)
  end if
end function

function EA_OnPostNewAttribute(Info)
 dim targetAttributeID 'Attribute in the BI-Group
 targetAttributeID = Info.Get("AttributeID")
 dim targetAttribute
 set targetAttribute = Repository.GetAttributeByID(targetAttributeID)
 dim targetElement
 set targetElement = Repository.GetElementByID(targetAttribute.ParentID)
 
 if targetElement.Stereotype = "REP_BI-Group" then
  targetAttribute.Stereotype = "REP_BI-Field"
  targetAttribute.Update
  addTaggedValue targetAttribute, "sourceAttribute", sourceAttribute.AttributeGUID
 end if
end function


function addTaggedValue(item, name, value)
 dim TVExist
 dim tv
 TVExist = false
 
 'first check if the tv exists
 for each tv in item.TaggedValues
  if tv.Name = name then
   TVExist = true
  end if
 next
 
 'if not, create the tv
 if not TVExist then
  set tv = item.TaggedValues.AddNew(name,"")
  tv.Value = value
  tv.Update
  item.Update
 else
  set tv = item.TaggedValues.GetByName(name)
  tv.Value = value
  tv.Update
  item.Update
 end if
 
 
 
end function