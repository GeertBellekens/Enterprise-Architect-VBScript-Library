'[path=\Projects\Project AC]
'[group=Acerta Scripts]
 
option explicit
 
!INC Local Scripts.EAConstants-VBScript
 
'
' Script Name:
' Author: Geert Bellekens
' Purpose: Sets all datatypes of columns to uppercase
' Date: 2016-1-10
'
sub main
       dim sqlUpdate
       'update attribute types
       sqlUpdate = "update a set a.type = upper(a.type) from t_attribute a where a.Stereotype = 'column'"
       Repository.Execute sqlUpdate
      
       'update parameter types
       sqlUpdate = "update opp set opp.Type = UPPER(opp.type) " & _
                           " from t_operationparams opp " & _
                           " inner join t_operation op on op.OperationID = opp.OperationID " & _
                           " inner join t_object o on o.Object_ID = op.Object_ID " & _
                           " where o.Stereotype = 'table' "
       Repository.Execute sqlUpdate
      
       'update connector roles for primary keys
       sqlUpdate = "update c set c.DestRole = op.Name, c.StyleEx = 'FKINFO=SRC=' + c.SourceRole + ':DST=' + op.Name + ':;' " & _
                           " from t_connector c " & _
                           " inner join t_object o on o.Object_ID = c.End_Object_ID " & _
                           " inner join t_operation op on op.Object_ID = o.Object_ID " & _
                                                                           " and op.Name like 'PK%' " & _
                                                                           " and op.Stereotype = 'PK' " & _
                           " where c.SourceRole like 'FK%' " & _
                           " and  " & _
                           " (isnull(c.DestRole,'') <>  op.Name " & _
                           " or " & _
                           " isnull(convert( varchar(500),c.StyleEx),'') <> 'FKINFO=SRC=' + c.SourceRole + ':DST=' + op.Name + ':;')"
       Repository.Execute sqlUpdate
end sub
 
main