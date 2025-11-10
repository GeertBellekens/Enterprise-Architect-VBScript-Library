'[path=\Framework\Utils]
'[group=Utils]


'
' Script Name: Translation
' Author: Geert Bellekens
' Purpose: Offers functions to deal with translations
' Date: 2025-08-27
'
const defaultLanguage = "en"


function createTranslatedCopy(diagram)
	dim secondaryLanguages
	set secondaryLanguages = getSecondaryLanguages()
	dim diagramOwner
	set diagramOwner = nothing
	if diagram.ParentID > 0 then
		set diagramOwner = Repository.GetElementByID(diagram.ParentID)
	end if
	if diagramOwner is nothing then
		set diagramOwner = Repository.GetPackageByID(diagram.PackageID)
	end if
	'get the language
	dim language
	language = getUserSelectedLanguage(secondaryLanguages)
	if len(language) = 0 then
		exit function
	end if
	dim translatedDiagram
	set translatedDiagram = copyDiagram(diagram, diagramOwner)
	'delete the previous translated diagram
	dim translatedDiagramName
	translatedDiagramName = diagram.Name & "_" & uCase(language)
	deleteExistingTranslateion diagram, translatedDiagramName
	'set the translated name
	translatedDiagram.Name = translatedDiagramName
	translatedDiagram.Update
	'set the "use alias switch
	if instr(translatedDiagram.ExtendedStyle, "UseAlias=0") > 0 then
		translatedDiagram.ExtendedStyle = replace(translatedDiagram.ExtendedStyle,"UseAlias=0","UseAlias=1")
	else
		translatedDiagram.ExtendedStyle = translatedDiagram.ExtendedStyle & "UseAlias=1;"
	end if
	translatedDiagram.Update
	translateDiagram translatedDiagram, language, true, secondaryLanguages
	'reload the original diagram
	Repository.ReloadDiagram diagram.DiagramID
	'open the translated diagram
	Repository.OpenDiagram translatedDiagram.DiagramID
	'reload the package of the diagram
	Repository.ReloadPackage diagram.PackageID
end function

function deleteExistingTranslateion(diagram, translatedDiagramName)
	'get owner of diagram
	dim owner as EA.Element
	set owner = nothing
	if diagram.ParentID > 0 then
		set owner = Repository.GetElementByID(diagram.ParentID)
	end if
	if owner is nothing then
		set owner = Repository.GetPackageByID(diagram.PackageID)
	end if
	dim i
	for i = owner.Diagrams.Count - 1 to 0 step -1
		dim currentDiagram as EA.Diagram
		set currentDiagram = owner.Diagrams.GetAt(i)
		if currentDiagram.Name = translatedDiagramName then
			owner.Diagrams.DeleteAt i, false
		end if
	next
end function

function translateItem(item, language, recursive, aliasOnly, secondaryLanguages)
	if item.ObjectType = otPackage then
		'translate the package object itself
		translateItem item.Element, language, recursive, aliasOnly, secondaryLanguages
		if recursive then
			'process elements
			dim element as EA.Element
			for each element in item.Elements
				translateElement element, language, recursive, aliasOnly, secondaryLanguages
			next
			'process subPackages
			dim subPackage as EA.Package
			for each subPackage in item.Packages
				translateItem subPackage, language, recursive, aliasOnly, secondaryLanguages
			next
		end if
	elseif item.ObjectType = otElement then
		'translate element
		translateElement item, language, recursive, aliasOnly, secondaryLanguages
	end if
end function

function translateElement(element, language, recursive, aliasOnly, secondaryLanguages)
	Repository.WriteOutput outPutName, now() & " Processing '" & element.Name & "'", 0
	dim dirty
	dirty = false
	if not aliasOnly then
		'name
		dim translatedName
		translatedName = element.GetTXName(language, 0)
		if len(translatedName) > 0 _
		  and not element.Name = translatedName then
			element.Name = translatedName
			dirty = true
		end if
		dim otherLanguage
		otherLanguage = getOtherLanguage(language, secondaryLanguages)
		'alias in the other language
		dim translatedAlias
		translatedAlias = element.GetTXName(otherLanguage, 0)
		if len(translatedAlias) > 0 _
		  and not element.Alias = translatedAlias then
			element.Alias = translatedAlias
			dirty = true
		end if
	else
		'fill alias with name
		translatedAlias = element.GetTXName(language, 0)
		if len(translatedAlias) > 0 _
		  and not element.Alias = translatedAlias then
			element.Alias = translatedAlias
			dirty = true
		end if
	end if
	'notes
	dim notes
	notes = getTranslatedNotes(element, secondaryLanguages)
	if not element.Notes = notes then
		element.Notes = notes
		dirty = true
	end if
	if dirty then
		element.update
	end if
	if recursive then
		dim subElement
		for each subElement in element.Elements
			translateElement subElement, language, recursive, aliasOnly, secondaryLanguages
		next
		for each subElement in element.EmbeddedElements
			translateElement subElement, language, recursive, aliasOnly, secondaryLanguages
		next
	end if
end function


function getTranslatedNotes(element, languages)
	dim notesText
	dim language
	'Name
	for each language in languages
		notesText = notesText &  "<b> Name " & Ucase(language) & ": </b>" & element.GetTXName(language, 0) & vbNewLine
	next
	'Notes
	'first check if the translated notes are empty. 
	if isTranslatedNotesEmpty(element, languages) then
	  if len(element.Notes) > 0 then
		''If they are, and the current notes are not, then we move the current notes to the translation for the default langauge
		element.SetTXNote defaultLanguage, element.Notes
	  else
		element.SetTXNote defaultLanguage, "<empty>"
	  end if
	end if
	
	for each language in languages
		notesText = notesText &  "<b> Notes " & Ucase(language) & ": </b>" & vbNewLine & element.GetTXNote(language, 0) & vbNewLine
	next
	'return
	getTranslatedNotes = notesText
end function

function isTranslatedNotesEmpty(element, languages)
	dim isEmpty
	isEmpty = true
	dim language
	for each language in languages
		dim translatedNotes
		translatedNotes = element.GetTXNote(language, 0)
		if len(translatedNotes) > 0 then
			isEmpty = false
		end if
	next
	'return
	isTranslatedNotesEmpty = isEmpty
end function

function translateProjectBrowser()
	dim secondaryLanguages 
	set secondaryLanguages = getSecondaryLanguages()
	dim language
	language = getUserSelectedLanguage(secondaryLanguages)
	if len(language) = 0 then
		exit function
	end if
	'reset output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	
	dim treeSelectedElements
	set treeSelectedElements = Repository.GetTreeSelectedElements()
	if treeSelectedElements.Count > 0 then
		'set timestamp
		Repository.WriteOutput outPutName, now() & " Starting " & outPutName & " for " & treeSelectedElements.Count & " selected elements" ,  0
		dim item
		for each item in treeSelectedElements
			translateItem item, language, true, false, secondaryLanguages
		next
		Repository.WriteOutput outPutName, now() & " Finished " & outPutName & " for " & treeSelectedElements.Count & " selected elements" ,  0
	else
	dim selectedItem
	set selectedItem = Repository.GetTreeSelectedObject
		Repository.WriteOutput outPutName, now() & " Starting " & outPutName & " for '"& selectedItem.Name &"'", 0
		translateItem selectedItem, language, true, false, secondaryLanguages
		Repository.WriteOutput outPutName, now() & " Finished " & outPutName & " for '"& selectedItem.Name &"'", 0
	end if
end function

function translateDiagram(diagram, language, aliasOnly,secondaryLanguages)
	'get the language
	if len(language) = 0 then
		language = getUserSelectedLanguage(secondaryLanguages)
	end if
	if len(language) = 0 then
		exit function
	end if
	'figure out if any element is selected
	dim elements
	set elements = getSelectedElementsOnDiagram(diagram)
	if elements.Count = 0 then
		set elements = getElementsOnDiagram(diagram)
	end if
	'translate the elements
	dim element as EA.Element
	for each element in elements
		translateItem element, language, false, aliasOnly, secondaryLanguages
	next
	'reload the diagram
	Repository.ReloadDiagram(diagram.DiagramID)
end function

function getUserSelectedLanguage(enabledLanguages)
	getUserSelectedLanguage = "" 'default empty string
	if enabledLanguages.Count = 0 then
		exit function
	end if
	'build inputbox string
	dim i
	i = 0
	dim selectMessage
	selectMessage = "Please enter the number for the corresponding language"
	dim language
	for each language in enabledLanguages
		i = i + 1
		selectMessage = selectMessage & vbNewLine & i & ": " & language
	next
	dim response
	response = InputBox(selectMessage, "Select Language", "1" )
	if isNumeric(response) then
		if Cstr(Cint(response)) = response then 'check if response is integer
			dim selectedID
			selectedID = Cint(response) - 1
			if selectedID >= 0 and selectedID < enabledLanguages.Count then
				'get the selected language
				language = enabledLanguages(selectedID)
			end if
		end if
	end if
	'return
	getUserSelectedLanguage = language
end function

function getOtherLanguage(language, secondaryLangauges)
	dim otherLanguage
	otherLanguage = ""
	dim currentLanguage
	for each currentLanguage in secondaryLangauges
		if not currentLanguage = language then
			otherLanguage = currentLanguage
			exit for
		end if
	next
	'return
	getOtherLanguage = otherLanguage
end function

function getSecondaryLanguages()
	dim secondaryLangauges
	set secondaryLangauges = CreateObject("System.Collections.ArrayList")
	'get all enabled languages
	dim sqlGetData
	sqlGetData = "select u.Value from usys_system u          " & vbNewLine & _
				" where u.Property = 'TranslateSecondary'   "
	dim languagesString
	languagesString = getSingleValueFromQuery(sqlGetData)
	if len(languagesString) > 0 then
		dim languageParts
		languageParts = split(languagesString, ";")
	
		dim languagePart
		for each languagePart in languageParts
			if len(languagePart) = 2 _
			  and not languagePart = defaultLanguage then 'language codes have two characters
				secondaryLangauges.Add languagePart
			end if
		next
	end if
	'return
	set getSecondaryLanguages = secondaryLangauges
end function