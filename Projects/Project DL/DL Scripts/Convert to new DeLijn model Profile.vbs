'[path=\Projects\Project DL\DL Scripts]
'[group=De Lijn Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Convert to new De Lijn model Profile
' Author: Geert Bellekens
' Purpose: Convert the CIM to the new profile
' Date: 2023-01-12
'



'name of the output tab
const outPutName = "Convert to DeLijn model profile"
dim stereotypeTranslations
dim beperkingsRegelTranslations
const translatedProfile = "Metamodel"

sub main
 'create output tab
 Repository.CreateOutputTab outPutName
 Repository.ClearOutput outPutName
 Repository.EnsureOutputVisible outPutName
 'create stereotype translations dictionary
 createTranslationsDictionary
 'get the selected package
 dim selectedPackage as EA.Package
 set selectedPackage = Repository.GetTreeSelectedPackage
 'make sure we are sure are you sure
 dim userIsSure
 userIsSure = Msgbox("Do you really want to converte package '" & selectedPackage.Name & "' ?", vbYesNo+vbQuestion, "Convert to DeLijn model profile?")
 if userIsSure = vbYes then
  'convert the selected package
  convertPackage selectedPackage
  'set timestamp for end
  Repository.WriteOutput outPutName,now() & " Finished Convert to DeLijn model profile"  , 0
 end if
end sub

function createTranslationsDictionary()
 set stereotypeTranslations = CreateObject("Scripting.Dictionary")
 stereotypeTranslations.Add "delijnproces"    ,"Business Proces"
 stereotypeTranslations.Add "archimate_businessrole"  ,"Business Rol"
 stereotypeTranslations.Add "cim_concept"    ,"Informatie Concept"
 stereotypeTranslations.Add "cim_begrip"     ,"Informatie Concept"
 stereotypeTranslations.Add "kpi"         ,"Indicator"
 
end function


function convertPackage(package)
 'set timestamp for start
 Repository.WriteOutput outPutName,now() & " Start converting package '" & package.Name & "'"  , 0
 'process elements
 dim element as EA.Element
 for each element in package.Elements
  'convert
  convertElement element
 next
 'process relations
 convertRelations(package)
 'process diagrams?
 'TODO
 'process subPackages
 dim subPackage as EA.Package
 for each subPackage in package.Packages
  convertPackage subPackage
 next
end function

function convertRelations(package)
 'get the relations we need using an SQL search
 dim sqlGetRelations
 sqlGetRelations = "select c.Connector_ID from t_connector c                              " & _
     " inner join t_object o on o.Object_ID = c.Start_Object_ID               " & _
     " where o.Package_ID = " & package.PackageID & "                         " 
 dim relations
 set relations = getConnectorsFromQuery(sqlGetRelations)
 dim relation as EA.Connector
 for each relation in relations
  dim translated
  translated = translateSingleStereotype(relation)
  if translated then
   dim sourceObject as EA.Element
   dim targetObject as EA.Element
   'get the source object
   set sourceObject = Repository.GetElementByID(relation.ClientID)
   'get the target object
   set targetObject = Repository.GetElementByID(relation.SupplierID)
   'inform user
   Repository.WriteOutput outPutName, now() & " Converted relation between '" & sourceObject.Name & "' and '" & targetObject.Name , sourceObject.ElementID
  end if
 next
end function

function convertElement(element)
 dim translated
 translated = translateSingleStereotype(element)
 if translated then
  'inform user
  Repository.WriteOutput outPutName, now() & " Converted element '" & element.Name & "'" , element.ElementID
 end if
end function


function setTaggedValue(element, tagName, tagValue)
 dim tv as EA.TaggedValue
 'refresh tagged values
 element.TaggedValues.Refresh
 'loop the tagged values
 for each tv in element.TaggedValues
  if lcase(tv.Name) = lcase(tagName) then
   tv.Value = tagValue
   tv.Update
   exit for
  end if
 next
end function

function translateSingleStereotype(item)
 'default false:
 translateSingleStereotype = false
 dim translatedStereotype
 'default empty string
 translatedStereotype = ""
 'check if stereotype is in list of translated stereotypes
 if stereotypeTranslations.Exists(lcase(item.Stereotype)) then
  'translate stereotype
  translatedStereotype = translatedProfile & "::" & stereotypeTranslations.Item(lcase(item.Stereotype))
  'keep a copy of the tagged values
  dim existingTags
  set existingTags = CreateObject("Scripting.Dictionary")
  dim tv as EA.TaggedValue
  for each tv in item.TaggedValues
   'exception: tag "Concept Type" should become stereotype "CIM_Begrip" if value is set to "Begrip"
   if tv.Name = "Concept Type" and tv.Value = "Begrip" then
    translatedStereotype = translatedProfile & "::CIM_Begrip"
   else
    existingTags.Add tv.Name, tv.Value
   end if
  next
  'set stereotypeEx
  item.StereotypeEx = translatedStereotype
  item.Update
  'reset tagged values
  dim tagName
  for each tagName in existingTags.Keys
   setTaggedValue item, tagName, existingTags(tagName)
  next
  'set return value to true
  translateSingleStereotype = true
 end if
end function




main