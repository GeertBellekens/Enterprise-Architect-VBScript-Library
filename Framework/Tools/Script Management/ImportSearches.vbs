'[path=\Framework\Tools\Script Management]
'[group=Script Management]


option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: ImportSearches
' Author: Geert Bellekens
' Purpose: Import searches from a folder containing search exports
' Date: 2018-08-10
'
const outPutName = "Import SQL Searches"

sub main
	'first get the searches file located at %appdata%\Sparx Systems\EA\Search Data\EA\EA_Search.xml
	dim shell,appData
	Set shell = CreateObject( "WScript.Shell" )
	appData = shell.ExpandEnvironmentStrings("%APPDATA%")
	dim eaSearchesFile
	set eaSearchesFile = new TextFile
	eaSearchesFile.FullPath = appData & "\Sparx Systems\EA\Search Data\EA_Search.xml"
	eaSearchesFile.loadContents
	Dim xSearchesDoc 
	Set xSearchesDoc = CreateObject( "MSXML2.DOMDocument" )
	If xSearchesDoc.LoadXML(eaSearchesFile.Contents) Then
		'then get the user selected folder and loop each file in it
		dim selectedFolder
		set selectedFolder = new FileSystemFolder
		set selectedFolder = selectedFolder.getUserSelectedFolder("")
		if not selectedFolder is nothing then
			'create output tab
			Repository.CreateOutputTab outPutName
			Repository.ClearOutput outPutName
			Repository.EnsureOutputVisible outPutName
			dim searchFile
			for each searchFile in selectedFolder.TextFiles
				dim xSearchDoc
				Set xSearchDoc = CreateObject( "MSXML2.DOMDocument" )
				if xSearchDoc.LoadXML(searchFile.Contents) then
					'now get the searches from the xSearchDoc and add or replace them into the xSearchesDoc
					addOrReplaceSearches xSearchesDoc, xSearchDoc
				end if
			next
			eaSearchesFile.Contents = xSearchesDoc.xml
			eaSearchesFile.Save
			Repository.WriteOutput outPutName, "Finished! Please restart EA for the searches to be reloaded", 0
		end if
	else
		msgbox "Please create at least one custom search"
	end if
	
end sub

function addOrReplaceSearches(xSearchesDoc,xSearchDoc)
	dim newSearchNodes, newSearchNode
	dim existingSearchesParentNode
	set existingSearchesParentNode = xSearchesDoc.SelectSingleNode("//RootSearch")
	set newSearchNodes = xSearchDoc.SelectNodes("//Search")
	if not existingSearchesParentNode is nothing then
		for each newSearchNode in newSearchNodes
			'get the GUID atrribute value
			dim guidAttribute
			set guidAttribute = newSearchNode.Attributes.GetNamedItem("GUID")
			'find the equivalent node in the xSearchesDoc and remove it
			dim equivalentSearchNode
			set equivalentSearchNode = xSearchesDoc.SelectSingleNode("//Search[@GUID='"& guidAttribute.Value &"']")
			if not equivalentSearchNode is nothing then
				Repository.WriteOutput outPutName, "Replacing search: " & newSearchNode.Attributes.GetNamedItem("Name").Value, 0
				existingSearchesParentNode.RemoveChild equivalentSearchNode 
			else
				Repository.WriteOutput outPutName, "Adding search: " & newSearchNode.Attributes.GetNamedItem("Name").Value, 0
			end if
			'add the new search
			existingSearchesParentNode.AppendChild newSearchNode
		next
	end if
end function

main