'[path=\Projects\Project Ar\Project Browser Package Group]
'[group=Project Browser Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Divide in packages per Diagram
' Author: Geert Bellekens
' Purpose: Puts each diagram in it's own subpackage
' Date: 2017-07-04
'
sub main
 dim package as EA.Package
 set package = Repository.GetTreeSelectedPackage()
 dim diagram as EA.Diagram
 for each diagram in package.Diagrams
  dim subPackage as EA.Package
  'create subPackage
  set subPackage = package.Packages.AddNew(diagram.Name,"Package")
  subPackage.Update
  'move diagram to subPackage
  diagram.PackageID= subPackage.PackageID
  diagram.Update
 next
 Repository.RefreshModelView package.PackageID
 msgbox "finished!"
end sub

main