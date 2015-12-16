'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit
'EA-Matic
'Author: Geert Bellekens
'This script will prevent any element to be deleted if it is still used as a type in either a parameter
'or an attribute. The can be overridden by first prepending the name with DELETED_

function EA_OnPreDeleteElement(Info)
     'Start by setting false
     EA_OnPreDeleteElement = false
     dim usage
     'Initialize the EAAddinFramework model
     dim model 
     set model = CreateObject("TSF.UmlToolingFramework.Wrappers.EA.Model")
     model.initialize(Repository)
     'get the elementID from Info
     dim elementID
     elementID = Info.Get("ElementID")
     'get the element being deleted
     dim element
     set element = model.getElementWrapperByID(elementID)     
     'Manual override is triggered by the name. If it starts with DELETED_ then the element may be deleted.
     if Left(element.name,LEN("DELETED_")) = "DELETED_" then
        'OK the element may be deleted
        EA_OnPreDeleteElement = true
     else
        dim usingAttributes
        set usingAttributes =  model.toArrayList(element.getUsingAttributes())
        'Check if the element is used as type in attributes
        if usingAttributes.Count = 0 then
            'Check if the element is used as type in a parameter
            dim usingParameters
            set usingParameters = model.toArrayList(element.getUsingParameters())
            if usingParameters.Count = 0 then
                'OK, no attributes or parameters use this element, it may be deleted
            EA_OnPreDeleteElement = true
            else
                usage = "parameter(s)"
            end if
        else
            usage = "attribute(s)"
        end if
     end if
     if EA_OnPredeleteElement = false then
          'NO the element cannot be deleted
          MsgBox "I'm sorry Dave, I'm afraid I can't do that" & vbNewLine _
          & element.name & " is used as type in " & usage , vbExclamation, "Cannot delete element"
     end if
end function