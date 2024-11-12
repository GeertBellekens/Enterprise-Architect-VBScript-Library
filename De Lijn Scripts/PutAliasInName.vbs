'[group=De Lijn Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: PutAliasInName
' Author: Geert Bellekens
' Purpose: add the alias to the name of all «DeLijnProcess» elements
' Date: 2022-10-07
'
const outPutName = "PutAliasInName"
const selectionStereotype = "DeLijnProces"
'const selectionStereotype = "ArchiMate_BusinessProcess"

sub main
 'create output tab
 Repository.CreateOutputTab outPutName
 Repository.ClearOutput outPutName
 Repository.EnsureOutputVisible outPutName
 dim package as EA.Package
 set package = Repository.GetTreeSelectedPackage
 'set timestamp for start
 Repository.WriteOutput outPutName,now() & " Start processing package '" & package.Name & "'"  , 0
 'do the actual work
 copyAliasToName package
 'set timestamp for end
 Repository.WriteOutput outPutName,now() & " Finished processing package '" & package.Name & "'"  , 0
end sub

function copyAliasToName(package)
 dim packageTreeIDString
 packageTreeIDString = getPackageTreeIDString(package)
 dim sqlGetData
 sqlGetData = "select o.Object_ID from t_object o                            " & vbNewLine & _
    " where o.Stereotype = '"& selectionStereotype &"'              " & vbNewLine & _
    " and len(o.alias) > 0                                          " & vbNewLine & _
    " and o.Package_ID in ("& packageTreeIDString &")               "
 dim elements
 set elements = getElementsFromQuery(sqlGetData)
 dim element as EA.Element
 dim i
 i = 0
 dim total
 total = elements.Count
 for each element in elements
  i = i + 1
  dim processElement
  processElement = false
  if len(element.Name) < len(element.Alias) then
   'if the element name is smaller than the alias then we certainly haven't copied it to the name yet
   processElement = true
  elseif not left(element.Name, len(element.Alias)) = element.Alias then
   'check if alias not already in the name
   processElement = true
  end if
  if processElement then
   Repository.WriteOutput outPutName,now() & " Updating element " & i & " of " & total & " : '" & element.Alias & " " & element.Name & "'"  , 0
   element.Name = element.Alias & " " & element.Name
   element.Update
  end if
 next
end function

main