'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit

!INC Local Scripts.EAConstants-VBScript

' EA-Matic
' Script Name: DefaultLineStyles
' Author: Geert Bellekens
' Purpose: Allows to determine the standard style of new connectors when they are created on a diagram
' Date: 14/03/2015
'
dim lsDirectMode, lsAutoRouteMode, lsCustomMode, lsTreeVerticalTree, lsTreeHorizontalTree, _
lsLateralHorizontalTree, lsLateralVerticalTree, lsOrthogonalSquareTree, lsOrthogonalRoundedTree

lsDirectMode = "1"
lsAutoRouteMode = "2"
lsCustomMode = "3"
lsTreeVerticalTree = "V"
lsTreeHorizontalTree = "H"
lsLateralHorizontalTree = "LH"
lsLateralVerticalTree = "LC"
lsOrthogonalSquareTree = "OS"
lsOrthogonalRoundedTree = "OR"

dim defaultStyle
dim menuDefaultLines


'*********EDIT BETWEEN HERE*************
' set here the menu name
menuDefaultLines = "&Set default linestyles"

' set here the default style to be used
defaultStyle = lsOrthogonalSquareTree

' set here the line style to be used for each type of connector
function determineLineStyle(connector)
	dim connectorType
	connectorType = connector.Type
	select case connectorType
		case "ControlFlow", "StateFlow","ObjectFlow","InformationFlow"
			determineLineStyle = lsOrthogonalRoundedTree
		case "Generalization", "Realization", "Realisation"
			determineLineStyle = lsTreeVerticalTree
		case "UseCase", "Dependency","NoteLink", "Abstraction"
			determineLineStyle = lsDirectMode
		case else
			determineLineStyle = defaultStyle
	end select
end function

'set here the color to be used for each type of connector
' use SparxColorFromRGB("E8", "8C", "0C") to get the correct integer color value
function determineColor(connector)
	' the default color
	determineColor = -1
end function

'set here the line width to be used for each type of connector
function determineLineWidth(connector)
	' the default line width
	determineLineWidth = 1
end function
'************AND HERE****************

'the event called by EA
function EA_OnPostNewConnector(Info)
	'get the connector id from the Info
	dim connectorID
	connectorID = Info.Get("ConnectorID")
	dim connector
	set connector = Repository.GetConnectorByID(connectorID)
	'get the current diagram
	dim diagram
	set diagram = Repository.GetCurrentDiagram()
	if not diagram is nothing then
		'first save the diagram
		Repository.SaveDiagram diagram.DiagramID
		'get the diagramlink for the connector
		dim diagramLink
		set diagramLink = getdiagramLinkForConnector(connector, diagram)
		if not diagramLink is nothing then
			'set the connectorstyle
			setConnectorStyle diagramLink, connector
			'save the diagramlink
			diagramLink.Update
			'reload the diagram to show the link style
			Repository.ReloadDiagram diagram.DiagramID
		end if
	end if
end function



'Tell EA what the menu options should be
function EA_GetMenuItems(MenuLocation, MenuName)
	if MenuName = "" and MenuLocation = "Diagram" then
		'Menu Header
		EA_GetMenuItems = menuDefaultLines
	end if
end function

' Do the work of setting default line styles
function DoDefaultLineStyles()
	dim diagram
	dim diagramLink
	dim connector
	dim dirty
	dirty = false
	set diagram = Repository.GetCurrentDiagram
	'save the diagram first
	Repository.SaveDiagram diagram.DiagramID
	'then loop all diagramLinks
	if not diagram is nothing then
		for each diagramLink in diagram.DiagramLinks
			set connector = Repository.GetConnectorByID(diagramLink.ConnectorID)
			if not connector is nothing then
				'set the connectorstyle
				setConnectorStyle diagramLink, connector
				'save the diagramlink
				diagramLink.Update
				dirty = true
			end if
		next
		'reload the diagram if we changed something
		if dirty then
			'reload the diagram to show the link style
			Repository.ReloadDiagram diagram.DiagramID
		end if
	end if
end function

'react to user clicking a menu option
function EA_MenuClick(MenuLocation, MenuName, ItemName)
	if ItemName = menuDefaultLines then
		DoDefaultLineStyles
	end if
end function

'gets the diagram link object
function getdiagramLinkForConnector(connector, diagram)
	dim diagramLink
	set getdiagramLinkForConnector = nothing
	for each diagramLink in diagram.DiagramLinks
		if diagramLink.ConnectorID = connector.ConnectorID then
			set getdiagramLinkForConnector = diagramLink
			exit for
		end if
	next
end function

'actually sets the connector style
function setConnectorStyle(diagramLink, connector)
	'split the style into its parts
	dim styleparts
	dim styleString
	' Throw away the last ; so that an empty cell at the end is not created when its Split
	if len(diagramLink.Style) > 0 then
		styleString = Left(diagramLink.Style, Len(diagramLink.Style)-1)
	else
		styleString = ""
	end if
	styleparts = Split(styleString,";")
	dim mode
	dim tree
	dim linestyle
	mode = ""
	tree = ""

	linestyle = determineLineStyle(connector)
	'these connectorstyles use mode=3 and the tree
	if  linestyle = lsTreeVerticalTree or _
		linestyle = lsTreeHorizontalTree or _
		linestyle = lsLateralHorizontalTree or _
		linestyle = lsLateralVerticalTree or _
		linestyle = lsOrthogonalSquareTree or _
		linestyle = lsOrthogonalRoundedTree then
		mode = "3"
		tree = linestyle
	else
		mode = linestyle
	end if
	'set the mode value
	setStylePart styleparts, "Mode", mode
	'set the tree value
	setStylePart styleparts, "TREE", tree

	setStylePart styleparts, "Color", determineColor(connector)
	setStylePart styleparts, "LWidth", determineLineWidth(connector)

	' update style (add in trailing ; that is needed)
	diagramLink.Style = join(styleparts, ";") & ";"
end function

' Set the style to the specified value
function setStylePart(styleparts, style, value)
	dim i
	dim stylePart
	dim index

	index = -1

	for i = 0 to Ubound(styleparts)
		stylePart = styleparts(i)
		if Instr(stylepart, style & "=") > 0 then
			index = i
		end if
	next


	If Len(value) > 0 then
		' Adding to style
		if index = -1 then
			' extend the array when style is not already in array
			redim preserve styleparts(Ubound(styleparts) + 1)
			index = Ubound(styleparts)
		end if
		styleparts(index) = style & "=" & value
	else
		' Removing style from styleparts
		if index >= 0 then
			' copy the last value over the top of index, and then shrink the array
			styleparts(index) = styleparts(Ubound(styleparts))
			redim preserve styleparts(Ubound(styleparts) - 1)
		end if
		' if the index was -1 it already did not exist in the styleparts
	end if

end function

' From http://www.sparxsystems.com.au/enterprise_architect_user_guide/11/automation_and_scripting/diagramobjects.html
' The color value is a decimal representation of the hex RGB value, where Red=FF, Green=FF00 and Blue=FF0000
' Who would write an RGB as BGR. YAEAB
function SparxColorFromRGB(red, green, blue)
	SparxColorFromRGB = CLng("&h" & blue & green & red)
end function