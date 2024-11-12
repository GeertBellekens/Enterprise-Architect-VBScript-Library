'[path=\Projects\Project A\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'



sub main
 dim sqlGetDSCalculations
 sqlGetDSCalculations = "select o.Object_ID from t_object o " & _
       " inner join t_objectproperties tv on tv.Object_ID = o.Object_ID " & _
       "        and tv.Property = 'SourceAttribute' " & _
       " where o.Stereotype = 'BI-DS_Calculation'"
 dim dsCalculations
 set dsCalculations = getElementsFromQuery(sqlGetDSCalculations)
 dim dsCalculation as EA.Element
 'loop the calculations and remove the tagged values
 for each dsCalculation in dsCalculations 
  dim i
  dim tv as EA.TaggedValue
  for i = dsCalculation.TaggedValues.Count -1 to 0 step -1
   dsCalculation.TaggedValues.DeleteAt i, false
   Session.Output "Removing tagged value from '" & dsCalculation.Name & "'"
  next
 next
end sub

main