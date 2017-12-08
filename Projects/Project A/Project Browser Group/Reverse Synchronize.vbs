'[path=\Projects\Project A\Project Browser Group]
'[group=Project Browser Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util

' Script Name: Reverse Synchronize
' Author: Geert Bellekens
' Purpose: Synchronises the names of the referencing elements/instances with this/called activity ref.
' Will also set the composite diagram to that of the classifier/ActivityRef in order to facilitate click-through
' Date: 19/03/2015
'


sub main
	' get the selected element
	dim counter
	counter = 0
	dim errors 
	errors = ""
	dim classifier as EA.Element
	set classifier = Repository.GetTreeSelectedObject
	if classifier.ObjectType = otElement then
		dim query
		query = ""
		dim instances
		set instances = nothing
		if classifier.Type = "Activity" then
			'synchronize calling activities
			query = "select o.Object_ID " &_
					"from (t_object o " &_
					"inner join t_objectproperties tv on tv.Object_ID = o.Object_ID) " &_
					"where tv.value = '" & classifier.ElementGUID & "' " &_
					" and tv.[Property] = 'calledActivityRef'"
		else
			'synchronize instances
			query = "select o.Object_ID from t_object o where o.Classifier > 0 and o.Classifier = " & classifier.ElementID
		end if
		if len(query) > 0 then
			set instances = Repository.GetElementSet(query, 2)
			dim instance as EA.Element
			for each instance in instances
				'apply userlock
				dim locked
				locked = instance.ApplyUserLock()
				'synchronize
				if locked then
					synchronizeElement instance
					'up the counter
					counter = counter +1
				else
					'element could not be synchronized
					if len(errors) = 0 then
						errors = "But could not synchronize: " & vbNewLine
					end if
					dim instancePackage as EA.Package
					set instancePackage = Repository.GetPackageByID(instance.PackageID)
					errors = errors & instancePackage.Name & "." & instance.Name & vbNewLine
				end if
			next
		end if
	end if
	dim message
	message = "Reverse Synchronize updated " & counter & " elements"
	if len(errors) > 0 then
		message = message & vbNewLine & errors
	end if
	MsgBox message
end sub

main