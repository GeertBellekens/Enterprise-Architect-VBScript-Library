'[path=\Projects\Project AC]
'[group=Acerta Scripts]
'
' Create a foreign key relationship between two database columns (attributes)
'
' @param[in] fkField (EA.Attribute) 
' @param[in] pkField (EA.Attribute) 
'
sub DefineForeignKey(fkField, pkField)
    dim pkOpName, fkOpName
    dim fkTable as EA.Element
    dim keyTable as EA.Element
    dim fkConnector as EA.Connector
    dim fkOperation as EA.Method
    dim op as EA.Method
    dim param as EA.Parameter

    set fkTable = Repository.GetElementByID(fkField.ParentID)
    set keyTable = Repository.GetElementByID(pkField.ParentID)
    
    ' get target Primary Key name
    for each op in keyTable.Methods
        if op.Stereotype = "PK" Then
            for each param in op.Parameters
                if param.Name = pkField.Name then
                    pkOpName = op.Name
                    exit for
                end if
            next
        end if
    next
    
    ' define Foreign Key Name
    fkOpName = "FK_" & fkTable.Name & "_" & keyTable.Name
    
    ' define connector
    Set fkConnector = fkTable.Connectors.AddNew("", "Association")
    fkConnector.SupplierID = pkField.ParentID
    fkConnector.StyleEx = "FKINFO=SRC=" & fkOpName & ":DST=" & pkOpName & ":;"
    fkConnector.StereotypeEx = "EAUML::FK"
    fkConnector.ClientEnd.Role = fkOpName
    fkConnector.ClientEnd.Cardinality = "0..*"
    fkConnector.SupplierEnd.Role = pkOpName
    fkConnector.SupplierEnd.Cardinality = "1"
    fkConnector.Update
    
    ' define fk operation
    set fkOperation = fkTable.Methods.AddNew(fkOpName, "")
    fkOperation.StereotypeEx = "EAUML::FK"
    fkOperation.Update
    
    set param = fkOperation.Parameters.AddNew(fkField.Name, fkField.Type)
    param.Update
    
    'set "On Delete" and "On Update" (optional)
    SetMethodTag fkOperation, "Delete", "Cascade"
    SetMethodTag fkOperation, "Update", "Set Null"
    SetMethodTag fkOperation, "property", "Delete Cascade=1;Update Set Null=1;"
	
    ' update attribute details
    fkField.IsCollection = true
    fkField.Update
    
end sub

function SetMethodTag(theMethod, tagName, tagValue)
    dim tag as EA.MethodTag
    set tag = theMethod.TaggedValues.GetByName(tagName)
    if tag is nothing then
        set tag = theMethod.TaggedValues.AddNew(tagName, "")
    end if
    tag.Value = tagValue
    tag.Update

    set SetMethodTag = tag
end function