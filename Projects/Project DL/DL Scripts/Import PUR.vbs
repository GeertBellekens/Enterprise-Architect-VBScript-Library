'[path=\Projects\Project DL\DL Scripts]
'[group=De Lijn Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Import Pur
' Author: Geert Bellekens
' Purpose: Import the data from the PUR exel into EA
' Date: 2018-09-19
'


'name of the output tab
const outPutName = "Import PUR"

sub main
 'create output tab
 Repository.CreateOutputTab outPutName
 Repository.ClearOutput outPutName
 Repository.EnsureOutputVisible outPutName
 'set timestamp for start
 Repository.WriteOutput outPutName,now() & " Start importing PUR"  , 0
 'open the excel file
 dim excel
 set excel = new ExcelFile
 excel.openUserSelectedFile
 'let the user select a base package to impor into
 msgbox "Please select the package to import into"
 dim importPackage as EA.Package
 set importPackage = selectPackage()

 if not importPackage is nothing then
  Repository.WriteOutput outPutName,now() & " Reading excel file"  , 0 
  'get the data from the "import" sheet
  dim purData
  purData = excel.getData("import")
  'actually import the data
  Repository.WriteOutput outPutName,now() & " Importing data"  , 0 
  importPurData purData, importPackage
 end if 
 'set timestamp for finish
 Repository.WriteOutput outPutName,now() & " Finished importing PUR"  , 0 
end sub

function importPurData(purData, importPackage)
 'loop the data
 dim i
 dim j
 'create the process dictionary 
 dim processDictionary
 set processDictionary = CreateObject("Scripting.Dictionary")
 for i = 5 to Ubound(purData)  'rows actual content starts at row 4
  dim ref, category, procesgroup, macroproces, process, description, comments, ownerDir, ownerFunction, ownerName, managerFunction, managerName, SD, IT, FI, HR, OP, TEC, SCM, MM, Ref2
  'reset values
  category = ""
  procesgroup = "" 
  macroproces = ""
  ref = ""
  process = ""
  description = ""
  comments = ""
  ownerDir = ""
  ownerFunction = ""
  ownerName = ""
  managerFunction = ""
  managerName = ""
  SD = ""
  IT = ""
  FI = ""
  HR = ""
  OP = ""
  TEC = ""
  SCM = ""
  MM = ""
  Ref2 = ""
  for j = 1 to Ubound(purData, 2)  'columns
   dim currentValue
   currentValue = purData(i,j)
   'fill in the values
   select case j
    case 1
     ref = CStr(currentValue) 'make sure it is a string value. If not it sometimes comes in as a double
     'if the reference ends with -00 then remove that part
     if right(ref, 3) = "-00" or right(ref, 3) = ".00"  then
      ref = left(ref, len(ref) -3)
     end if
     'replace "." with "-"
     ref = trim(replace(ref, ".", "-"))
     'add a leading zero if needed
     dim refParts
     refParts = split(ref, "-")
     dim k
     for k = 0 to Ubound(refParts)
      if len(refParts(k)) < 3 then
       'add leading zero
       refParts(k) = "0" + refParts(k)
      end if
     next
     ref = Join(refParts,"-")
    case 2
     category = currentValue
    case 3
     procesgroup = currentValue
    case 4
     macroproces = currentValue
    case 5
     process = currentValue
    case 6
     description = currentValue
    case 7
     comments = currentValue
    case 8
     ownerDir = currentValue
    case 9
     ownerFunction = currentValue
    case 10
     ownerName = currentValue
    case 11
     managerFunction = currentValue
    case 12
     managerName = currentValue
    case 13
     SD = currentValue
    case 14
     IT = currentValue
    case 15
     FI = currentValue
    case 16
     HR = currentValue
    case 17
     OP = currentValue
    case 18
     TEC = currentValue
    case 19
     SCM = currentValue
    case 20
     MM = currentValue
    case 21
     Ref2 = currentValue
   end select
   'debug
   'Repository.WriteOutput outPutName,now() & "Field(" & i & "," & j & ") : " & currentValue, 0
  next
  'add the process only if the ref is filled in
  if len(ref) > 0 then
   addProcess processDictionary, importPackage, ref, category, procesgroup, macroproces, process, description, comments, ownerDir, ownerFunction, ownerName, managerFunction, managerName, SD, IT, FI, HR, OP, TEC, SCM, MM, Ref2
  end if
 next
end function

function addProcess( processDictionary, importPackage, ref, category, procesgroup, macroproces, process, description, comments, ownerDir, ownerFunction, ownerName, managerFunction, managerName, SD, IT, FI, HR, OP, TEC, SCM, MM, Ref2)
 'get the name of the current element to be added
 dim name
 name = category
 if len(procesgroup) > 0 then
  name = procesgroup
 end if
 if len(macroproces) > 0 then
  name = macroproces
 end if
  if len(process) > 0 then
  name = process
 end if
 if len(name) = 0 then
  'report error
  Repository.WriteOutput outPutName,now() & " ERROR: Name empty for reference: " & ref, 0
  exit function
 end if
 'check if already in the dictionary
 if processDictionary.Exists(ref) then
  set prElement = processDictionary(ref)
 else
  'check if we already have an object with the given ref
  dim prElement as EA.Element
'  set prElement = getExistingProcessElement(name)
'  if not prElement is nothing then
'   'move to the import package
'   prElement.PackageID = importPackage.PackageID
'   prElement.Update
'  else
   set prElement = createNewProcessElement(name, importPackage)
'  end if
  'add the element to the dictionary
  processDictionary.Add ref, prElement
 end if
 'set the properties
 'description
 prElement.Notes = description
 'reference
 prElement.Alias = ref
 'comments
 dim commentTag as EA.TaggedValue
 set commentTag = getOrCreateTaggedValue(prElement,"comments")
 commentTag.Value = "<memo>"
 commentTag.Notes = comments
 commentTag.Update
 'PUR 2.0
' if len(ref2) > 0 then
'  dim pur2Tag as EA.TaggedValue
'  set pur2Tag = getOrCreateTaggedValue(prElement,"PUR 2.0")
'  pur2Tag.Value = Ref2
'  pur2Tag.Update
' end if
 if name = macroproces or name = procesgroup  then 'save owner and raakvlakken only on macro process level
  'owner
  dim owner
  owner = getOwner(ownerDir, SD, IT, FI, HR, OP, TEC, SCM, MM)
  createTaggedValueWithValue prElement, "eigenaar directie", owner
  'owner function
  createTaggedValueWithValue prElement, "eigenaar functie", ownerFunction
  'owner name
  createTaggedValueWithValue prElement, "eigenaar naam", ownerName
  'manager function
  createTaggedValueWithValue prElement, "beheerder functie", managerFunction
  'manager naam
  createTaggedValueWithValue prElement, "beheerder naam", managerName
  'raakvlakken
  dim raakvlakken
  raakvlakken = getRaakVlakken(SD, IT, FI, HR, OP, TEC, SCM, MM)
  createTaggedValueWithValue prElement, "raakvlakken", raakvlakken
 end if
 'save the element
 prElement.Update
 'link to parent
 linkToParent prElement, processDictionary
end function

function createTaggedValueWithValue(element, tagname, tagValue)
 dim tag as EA.TaggedValue
 set tag = getOrCreateTaggedValue(element,tagname)
 tag.Value = tagValue
 tag.Update
end function

function getRaakVlakken(SD, IT, FI, HR, OP, TEC, SCM, MM)
 dim raakvlakList
 set raakvlakList = CreateObject("System.Collections.ArrayList")
 if Ucase(SD) = "X" then
  raakvlakList.Add "SD"
 elseif Ucase(IT) = "X" then
  raakvlakList.Add "IT"
 elseif Ucase(FI) = "X" then
  raakvlakList.Add "FI"
 elseif Ucase(HR) = "X" then
  raakvlakList.Add "HR"
 elseif Ucase(OP) = "X" then
  raakvlakList.Add "OP"
 elseif Ucase(TEC) = "X" then
  raakvlakList.Add "TEC"
 elseif Ucase(MM) = "X" then
  raakvlakList.Add "MM"
 end if
 'convert to array
 dim raakVlakArray
 raakVlakArray = raakvlakList.ToArray()
 'join as string and return
 getRaakVlakken = Join(raakVlakArray,",")
end function

function getOwner(ownerDir, SD, IT, FI, HR, OP, TEC, SCM, MM)
 dim owner
 owner = ""
 if Ucase(SD) = "L" then
  owner = "SD"
 elseif Ucase(IT) = "L" then
  owner = "IT"
 elseif Ucase(FI) = "L" then
  owner = "FI"
 elseif Ucase(HR) = "L" then
  owner = "HR"
 elseif Ucase(OP) = "L" then
  owner = "OP"
 elseif Ucase(TEC) = "L" then
  owner = "TEC"
 elseif Ucase(MM) = "L" then
  owner = "MM"
 else
  owner = ownerDir
 end if
 'return 
 getOwner = owner
end function

function linkToParent(prElement, processDictionary)
 'get the ref
 dim parentRef 
 parentRef = ""
 'trim off the last .xx
 dim dotLoc
 dotLoc = InStrRev (prElement.Alias,"-")
 if dotLoc > 1 then
  parentRef = left(prElement.Alias, dotLoc -1)
 end if
 dim parentPackageID
 parentPackageID = prElement.PackageID 'initialize
 if len(parentRef) > 0 then
  'get the parent
  dim parentProcess as EA.Element
  set parentProcess = nothing
  if processDictionary.Exists(parentRef) then
   set parentProcess = processDictionary(parentRef)
  end if
  if not parentProcess is nothing then
   'get the parent package id
   parentPackageID = parentProcess.PackageID
   'check if not already present
   dim existingCompositions
   dim sqlGetExistingCompositions
   sqlGetExistingCompositions = "select c.Connector_ID from t_connector c     " & _
          " where c.Stereotype = 'ArchiMate_Composition' " & _
          " and c.End_Object_ID = " & parentProcess.ElementID &  _
          " and c.Start_Object_ID = " & prElement.ElementID
   set existingCompositions = getConnectorsFromQuery(sqlGetExistingCompositions)
   if existingCompositions.Count = 0 then
    'create Archimate composition to parent
    dim composition as EA.Connector
    set composition = prElement.Connectors.AddNew("","Archimate2::ArchiMate_Composition")
    composition.SupplierID = parentProcess.ElementID
    composition.SupplierEnd.Aggregation = 2 'Composite
    'set direction
    composition.Direction = "Source -> Destination"
    composition.ClientEnd.Navigable = "Unspecified"
    composition.ClientEnd.Update
    composition.SupplierEnd.Navigable = "Navigable"
    composition.SupplierEnd.Update
    'save
    composition.Update
   end if
  else
   'report error
   Repository.WriteOutput outPutName,now() & " Could not find parent process with reference: " & parentRef, prElement.ElementID
   'Debug print out content of the process dictionary
   dim key
   for each key in processDictionary.Keys
    Repository.WriteOutput outPutName,now() & "Key: " & key & "Process: " & processDictionary(key).Name, 0
   next
  end if
 end if
 'create the package and move the element
 createPackageForProcess prElement, parentPackageID
end function

function createPackageForProcess (prElement, parentPackageID)
 'check level
 dim refParts
 refParts = Split(prElement.Alias, "-")
 dim packageName
 if Ubound(refParts) = 0 or Ubound(refParts) = 1 then
  packageName = prElement.Alias & " " & prElement.Name
 elseif Ubound(refParts) = 2 then
  packageName =  prElement.Name
 elseif Ubound (refParts) = 3 then
  packageName = "Processes"
 else
  packageName = ""
 end if
 if len(packageName) > 0 then
  'find existing package
  dim parentPackage as EA.Package
  set parentPackage = Repository.GetPackageByID(parentPackageID)
  dim package as EA.Package
  dim processPackage as EA.Package
  set processPackage = nothing
  for each package in parentPackage.Packages
   if package.Name = packageName then
    set processPackage = package
    exit for
   end if
  next
  'create new package if not found
  if processPackage is nothing then
   set processPackage = parentPackage.Packages.AddNew(packageName, "")
   processPackage.Update
  end if
  'move process to this new package
  prElement.PackageID = processPackage.PackageID
  prElement.update
 end if
end function

function createNewProcessElement(name, importPackage)
 dim newElement as EA.Element
 set newElement = importPackage.Elements.AddNew(name, "Archimate2::ArchiMate_BusinessProcess")
 newElement.Update
 set createNewProcessElement = newElement
 'debug
 Repository.WriteOutput outPutName,now() & " Created new element: " & name, newElement.ElementID
end function

function getExistingProcessElement(name)
 dim prElement as EA.Element
 set prElement = nothing 'initialize
 dim sqlGetProcess
 sqlGetProcess = " select o.Object_ID from t_object o " & _
     " where o.Stereotype = 'ArchiMate_BusinessProcess' " & _
     " and o.Name = '" & replace(name, "'","''") & "' "
 dim prElements
 set prElements = getElementsFromQuery(sqlGetProcess)
 if prElements.Count = 1 then
  set prElement = prElements(0) 'we only need the first one
  'debug
  Repository.WriteOutput outPutName,now() & " Found existing element: " & prElement.Name, prElement.ElementID
 elseif prElements.Count > 1 then
  'warn if multiple found for the same reference
  Repository.WriteOutput outPutName,now() & " ERROR: found multiple elements for name: " & Name , 0
 end if
 'return
 set getExistingProcessElement = prElement
end function 

main