'[path=\Projects\Project A\Project Browser Package Group]
'[group=Project Browser Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

const outPutName = "Get Package ID List"

sub main
  Repository.CreateOutputTab outPutName
  Repository.ClearOutput outPutName
  Repository.EnsureOutputVisible outPutName
  
  dim selectedPackage as EA.Package
  set selectedPackage = Repository.GetTreeSelectedPackage
  dim packageTree
  set packageTree = getPackageTree(selectedPackage)
  dim packageIDList
  packageIDList = makePackageIDString(packageTree)
  'set timestamp
  Repository.WriteOutput outPutName, "PackageIDString for package " & selectedPackage.Name & " : " & packageIDList,0
end sub

main