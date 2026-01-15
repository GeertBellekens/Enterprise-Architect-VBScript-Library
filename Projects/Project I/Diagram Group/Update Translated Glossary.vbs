'[path=\Projects\Project I\Diagram Group]
'[group=Diagram Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' Script Name: Update Translated Glossary
' Author: Geert Bellekens
' Purpose: For each element on this glossary diagram, create translated entries into the glossary
' Date: 2025-11-19

const outPutName = "Update Translated Glossary"


sub main
	'reset output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName

	dim diagram as EA.Diagram
	
	set diagram = Repository.GetCurrentDiagram
	if diagram is nothing then
		exit sub
	end if
	'check diagram type
	if not diagram.MetaType = "Glossary Item Lists::GlossaryItemList" then
		msgbox "This script only works on GlossaryItemList diagrams"
		exit sub
	end if
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName & " for '"& diagram.Name &"'", 0
	'do the actual work
	dim secundaryLanguages
	set secundaryLanguages = getSecondaryLanguages()
	updateTranslatedGlossary diagram, secundaryLanguages
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName & " for '"& diagram.Name &"'", 0
	
end sub

function updateTranslatedGlossary(diagram, secundaryLanguages)
	'save diagram before continuing
	Repository.SaveDiagram diagram.DiagramID
	'get all elements on the diagram
	dim elements
	set elements = getElementsOnDiagram(diagram)
	dim language
	for each language in secundaryLanguages
		dim category
		category = diagram.Name & "." & ucase(language)
		'get all glossary entries for this diagram and language
		dim glossaryEntries
		set glossaryEntries = getGlossaryEntries(category)
		'loop all elements
		dim element as EA.Element
		for each element in elements
			createOrUpdateTranslatedTerm element,category,language,glossaryEntries
		next
	next
end function

function createOrUpdateTranslatedTerm(element,category,language,glossaryEntries )
	dim translatedTerm
	translatedTerm = element.GetTXName(language, 0)
	if len(translatedTerm) > 0 then
		dim translatedMeaning
		translatedMeaning = element.GetTXNote(language, 0)
		dim term as EA.Term
		if not glossaryEntries.Exists(lcase(translatedTerm)) then
			'report progress
			Repository.WriteOutput outPutName, now() & " Creating term '" & translatedTerm & "'", 0
			set term = Repository.Terms.AddNew(translatedTerm, category)
			term.Meaning = translatedMeaning
			term.Update
		else
			set term = glossaryEntries(lcase(translatedTerm))
		end if
		if not term.Meaning = translatedMeaning then
			'report progress
			Repository.WriteOutput outPutName, now() & " Updating term '" & translatedTerm & "'", 0
			term.Meaning = translatedMeaning
			term.Update
		end if
	end if
end function

function getGlossaryEntries(category)
	dim term as EA.Term
	dim glossaryEntries
	set glossaryEntries = CreateObject("Scripting.Dictionary")
	'refresh glossary to make sure we have the correct set
	Repository.Terms.Refresh
	for each term in Repository.Terms
		if lcase(term.Type) = lcase(category) then
			if not glossaryEntries.Exists(lcase(term.Term)) then
				glossaryEntries.Add lcase(term.Term), term
			end if
		end if
	next
	'return
	set getGlossaryEntries = glossaryEntries
end function

main