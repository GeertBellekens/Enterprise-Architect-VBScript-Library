'[path=\Projects\Project DL\DL Scripts]
'[group=De Lijn Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Convert CIM
' Author: Geert Bellekens
' Purpose: Convert the CIM to the new profile
' Date: 2018-09-05
'

'TODO: CRUD relaties zijn van Process naar Concept, dus query voor converteren relatie moet hier rekening mee houden.


'name of the output tab
const outPutName = "Convert CIM"
dim stereotypeTranslations
dim beperkingsRegelTranslations
const translatedProfile = "CM"

sub main
 'create output tab
 Repository.CreateOutputTab outPutName
 Repository.ClearOutput outPutName
 Repository.EnsureOutputVisible outPutName
 'set timestamp for start
 Repository.WriteOutput outPutName,now() & " Start converting CIM"  , 0
 'create stereotype translations dictionary
 createTranslationsDictionary
 'and the BeperkingsRegel translations
 createBeperkingsRegelTranslations
 'get the selected package
 dim selectedPackage as EA.Package
 set selectedPackage = Repository.GetTreeSelectedPackage
 'convert the selected package
 convertCIM selectedPackage
 'set timestamp for end
 Repository.WriteOutput outPutName,now() & " Finished converting CIM"  , 0
end sub


function convertCIM(package)
 'process elements
 dim element as EA.Element
 for each element in package.Elements
  'convert
  convertElement element
 next
 'process relations
 convertRelations(package)
 'process subPackages
 dim subPackage as EA.Package
 for each subPackage in package.Packages
  convertCIM subPackage
 next
end function

function convertRelations(package)
 'get the relations we need using an SQL search
 dim sqlGetRelations
 sqlGetRelations = "select c.Connector_ID from t_connector c                              " & _
     " inner join t_object o on o.Object_ID = c.Start_Object_ID               " & _
     " where o.Package_ID = " & package.PackageID & "                         " & _
     " union                                                                  " & _
     " select c.Connector_ID from t_connector c                               " & _
     " inner join t_object o on o.Object_ID = c.start_object_id               " & _
     "                         and o.stereotype = 'ArchiMate_BusinessProcess' " & _
     " inner join t_object o2 on o2.Object_ID = c.end_object_id               " & _
     " where o2.package_ID = " & package.PackageID
 dim relations
 set relations = getConnectorsFromQuery(sqlGetRelations)
 dim relation as EA.Connector
 for each relation in relations
  dim translated
  translated = translateSingleStereotype(relation)
  
  dim sourceObject as EA.Element
  dim targetObject as EA.Element
  if translated then
   'get the source object
   set sourceObject = Repository.GetElementByID(relation.ClientID)
   'get the target object
   set targetObject = Repository.GetElementByID(relation.SupplierID)
   'inform user
   Repository.WriteOutput outPutName, now() & " Converted relation between '" & sourceObject.Name & "' and '" & targetObject.Name , sourceObject.ElementID
   'reverse the roles in case of an association
   if relation.Type = "Association" then
    'inform user
    Repository.WriteOutput outPutName,now() & " Reversing '" & relation.ClientEnd.Role & "' with '" & relation.SupplierEnd.Role & _
           "' for association with GUID " & relation.ConnectorGUID  , relation.ClientID
    'reverse the roles
    dim tempRole
    tempRole = relation.ClientEnd.Role
    relation.ClientEnd.Role = relation.SupplierEnd.Role
    relation.SupplierEnd.Role = tempRole
    'relation.Color = 10058240 'blue
    relation.Update
   end if
  end if
  'if it is a dependency starting at a BeperkingsRegel then we add the stereotype CIM_Beperking
  'if it is a dependency starting at an Archimate BusinessProcess and going to a CIM_Concept then we
  ' - add stereotype CIM_CRUD
  ' - set the tagged value based on the targetRole
  ' - clear source and target role
  if relation.Type = "Dependency" or relation.Type = "Abstraction" then
   'get the source object
   set sourceObject = Repository.GetElementByID(relation.ClientID)
   'get the target object
   set targetObject = Repository.GetElementByID(relation.SupplierID)
   if instr(sourceObject.Stereotype, "Beperkingsregel") > 0_
     and instr(targetObject.Stereotype, "Concept") > 0 then
    'set the stereotype
    relation.StereotypeEx = translatedProfile & "::CIM_Beperking"
    relation.Type = "Dependency"
    relation.Update
    'inform user
    Repository.WriteOutput outPutName, now() & " Converted relation between '" & sourceObject.Name & "' and '" & targetObject.Name , sourceObject.ElementID
   elseif sourceObject.Stereotype = "ArchiMate_BusinessProcess" _
     and instr(targetObject.Stereotype, "Concept") > 0 then
    'set the stereotype
    relation.StereotypeEx = translatedProfile & "::CIM_CRUD"
    relation.Type = "Dependency"
    relation.Update
    if relation.SupplierEnd.Role = "C" _
      or relation.SupplierEnd.Role = "R" _
      or relation.SupplierEnd.Role = "U" _
      or relation.SupplierEnd.Role = "D" then
     'set the tagged value based on the targetRole
     setTaggedValue relation, "CRUD", relation.SupplierEnd.Role
     'clear source and target role
     relation.SupplierEnd.Role = ""
     relation.SupplierEnd.Update
     relation.ClientEnd.Role = ""
     relation.ClientEnd.Update
    end if
    'inform user
    Repository.WriteOutput outPutName, now() & " Converted relation between '" & sourceObject.Name & "' and '" & targetObject.Name , sourceObject.ElementID
   end if
  end if
  'set line color to dark blue => 10058240
  if left(relation.Stereotype, 4) = "CIM_" then
   relation.Color = 10058240 
   relation.Update
  end if
 next
end function

function convertElement(element)
 dim translated
 translated = translateSingleStereotype(element)
 if translated then
  setBeperkingsRegelType element
  'inform user
  Repository.WriteOutput outPutName, now() & " Converted element '" & element.Name & "'" , element.ElementID
 end if
end function

function setBeperkingsRegelType(element)
 if element.Stereotype = "CIM_Beperkingsregel" then
  dim keyword
  for each keyword in beperkingsRegelTranslations.Keys
   if instr(element.Name, keyword) > 0 then
    setTaggedValue element, "Constraint Type", beperkingsRegelTranslations(keyword)
    'return 
    exit function
   end if
  next
  'if we get here then we set the default to "Procedureel"
  setTaggedValue element, "Constraint Type", "Procedureel"
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

function createTranslationsDictionary()
 set stereotypeTranslations = CreateObject("Scripting.Dictionary")
 stereotypeTranslations.Add "cim concept"    ,"CIM_Concept"
 stereotypeTranslations.Add "zin"      ,"CIM_Zin"
 stereotypeTranslations.Add "deelverzameling"   ,"CIM_Deelverzameling"
 stereotypeTranslations.Add "beperkingsregel"   ,"CIM_Beperkingsregel"
end function

function createBeperkingsRegelTranslations()
 set beperkingsRegelTranslations = CreateObject("Scripting.Dictionary")
 beperkingsRegelTranslations.Add ":Unic"  ,"Uniek"
 beperkingsRegelTranslations.Add ":Subtype" ,"Subtype"
 beperkingsRegelTranslations.Add ":Tota"  ,"Totaliteit"
 beperkingsRegelTranslations.Add ":Gelijk" ,"Gelijkheid"
 beperkingsRegelTranslations.Add ":Exclu" ,"Exclusiviteit"
end function

main