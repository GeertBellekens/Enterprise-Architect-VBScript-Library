'[path=\Projects\Project DL\DL Scripts]
'[group=De Lijn Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Reverse association Roles 
' Author: Geert Bellekens
' Purpose: Put the name of the source rol at the target role and put the name of the target role at the source role.
' Date: 2018-09-05
'

'name of the output tab
const outPutName = "Reverse association roles"

sub main
 'create output tab
 Repository.CreateOutputTab outPutName
 Repository.ClearOutput outPutName
 Repository.EnsureOutputVisible outPutName
 'set timestamp for start
 Repository.WriteOutput outPutName,now() & " Starting Reverse Association roles"  , 0
 'get the selected package
 dim selectedPackage as EA.Package
 set selectedPackage = Repository.GetTreeSelectedPackage
 dim packageTreeIDs
 packageTreeIDs = getPackageTreeIDString(selectedPackage)
 'get associations of the current diagram
 dim sqlGetAssociations
 sqlGetAssociations = "select c.Connector_ID from t_connector c                  " & _
      " inner join t_object o on c.Start_Object_ID = o.Object_ID  " & _
      " where c.Connector_Type =  'Association'                   " & _
      " and o.Package_ID in (" & packageTreeIDs &")               "
 dim associations 
 set associations = getConnectorsFromQuery(sqlGetAssociations)
 dim association as EA.Connector
 for each association in associations
  'inform user
  Repository.WriteOutput outPutName,now() & " Reversing '" & association.ClientEnd.Role & "' with '" & association.SupplierEnd.Role & _
         "' for association with GUID " & association.ConnectorGUID  , association.ClientID
  'actually reverse the roles
  dim tempRole
  tempRole = association.ClientEnd.Role
  association.ClientEnd.Role = association.SupplierEnd.Role
  association.SupplierEnd.Role = tempRole
  association.Update
 next
 'inform user
 Repository.WriteOutput outPutName,now() & " Finished Reverse Association roles"  , 0
end sub

main