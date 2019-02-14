'[group=Temp]
Sub MigrateElement (sGUID, lngPackageID)

     Dim proj as EA.Project

     set proj = Repository.GetProjectInterface

     proj.Migrate sGUID, "BPMN1.1", "BPMN2.0"

     'refresh the model

     If lngPackageID<>0 Then

          Repository.RefreshModelView (lngPackageID)

     End If

End Sub

Sub MigrateSelectedItem

     Dim selType

     Dim selElement as EA.Element

     Dim selPackage as EA.Package

     selType = GetTreeSelectedItemType

     If selType = 4 Then 'means Element

          set selElement = GetTreeSelectedObject

          MigrateElement selElement.ElementGUID, selElement.PackageID

          MsgBox "Element Migration Completed",0,"BPMN 2.0 Migration"

     ElseIf selType = 5 Then 'means Package

          set selPackage = GetTreeSelectedObject

          MigrateElement selPackage.PackageGUID, selPackage.PackageID

          MsgBox "Package Migration Completed",0,"BPMN 2.0 Migration"

     Else

          MsgBox "Select a Package or Element in the Project Browser to initiate migration",0,"BPMN 2.0 Migration"

     End If

End Sub

Sub Main

     MigrateSelectedItem

End Sub

Main