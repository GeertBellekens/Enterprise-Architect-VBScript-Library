'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

const outPutName = "Get Package ID List"

sub main
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		Repository.WriteOutput outPutName, getCurrentPackageTreeIDString(),0
end sub

main