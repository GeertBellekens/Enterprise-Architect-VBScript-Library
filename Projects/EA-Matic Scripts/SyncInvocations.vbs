'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit

!INC Local Scripts.EAConstants-VBScript

' EA-Matic
' Author: Geert Belleekns
' Purpose: This EA-Matic script will keep the name of objects the same as the name of their classifier
' More info: 

function EA_OnNotifyContextItemModified(GUID, ot)
	dim model
	set model = getEAAddingFrameworkModel()
	'only do something when the changed object is an element
	if ot = otElement then
		dim element
		set element = model.getElementWrapperByGUID(GUID)
		synchronizeObjectNames element, model
	end if
end function

function EA_OnPostNewElement(Info)
	'Get the model
	dim model
	set model = getEAAddingFrameworkModel()
	'get the elementID from Info
    dim elementID
    elementID = Info.Get("ElementID")
    'get the element being deleted
    dim element
    set element = model.getElementWrapperByID(elementID)
	synchronizeObjectNames element, model
end function

'gets a new instance of the EAAddinFramework and initializes it with the EA.Repository
function getEAAddingFrameworkModel()
	'Initialize the EAAddinFramework model
    dim model 
    set model = CreateObject("TSF.UmlToolingFramework.Wrappers.EA.Model")
    model.initialize(Repository)
	set getEAAddingFrameworkModel = model
end function

function synchronizeObjectNames(element, model)
	'first check if this is an object
	if element.WrappedElement.Type = "Action" AND element.WrappedElement.ClassifierID > 0 then
		dim classifier
		set classifier = model.getElementWrapperByID(element.WrappedElement.ClassifierID)
		if not classifier is nothing AND classifier.name <> element.name then
			element.name = classifier.name
			element.save
		end if
	else
		'get all objects having this element as their classifier
		dim query
		query = "select o.Object_ID from t_object o where o.classifier =" & element.id
		dim objects
		set objects = model.toArrayList(model.getElementWrappersByQuery(query))
		'loop objects
		dim obj
		for each obj in objects
			'rename the object if the name is different from the classifiers name
			if obj.name <> element.name then
				obj.name = element.name
				obj.save
			end if 
		next
	end if
end function