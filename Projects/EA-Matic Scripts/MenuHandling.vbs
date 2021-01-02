'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
'EA-Matic
!INC Local Scripts.EAConstants-VBScript

'Author: Geert Bellekens
'This script demonstrates how to add menu options to the add-in menu and how to react on a menu click.
'
''Tell EA what the menu options should be
'function EA_GetMenuItems(MenuLocation, MenuName)
'	if MenuName = "" then
'		'Menu Header
'		EA_GetMenuItems = "-&MyAddinMenu"
'	else 
'		if MenuName = "-&MyAddinMenu" then
'			'Menu items
'			Dim menuItems(1)
'			 menuItems(0) = "TreeViewMenu"
'			 menuItems(1) = "DiagramMenu"
'			 EA_GetMenuItems = menuItems 
'		 end if
'	end if 
'end function
'
''Define the state of the menu options
'function EA_GetMenuState(MenuLocation, MenuName, ItemName, IsEnabled, IsChecked)
'	if MenuName = "-&MyAddinMenu" then
'		Select Case ItemName
'			case "TreeViewMenu"
'				if MenuLocation = "TreeView" then
'					IsEnabled = true
'				else
'					IsEnabled = false
'				end if
'			case "DiagramMenu"
'				if MenuLocation = "Diagram" then
'					IsEnabled = true
'				else
'					IsEnabled = false
'				end if
'		end select
'	end if
'	'to return out parameter values we should return an array with all parameters
'	EA_GetMenuState = Array(MenuLocation, MenuName, ItemName, IsEnabled, IsChecked)
'end function
'
''react to user clicking a menu option
'function EA_MenuClick(MenuLocation, MenuName, ItemName)
'	 	if MenuName = "-&MyAddinMenu" then
'		Select Case ItemName
'			case "TreeViewMenu"
'				Dim Package
'				Set Package = Repository.GetTreeSelectedPackage()
'				MsgBox ("Current Package is: " & Package.Name)
'			case "DiagramMenu"
'				Dim Diagram
'				Set Diagram = Repository.GetCurrentDiagram()
'				MsgBox("Current Diagram is: " & Diagram.Name)
'		end select
'	end if
'end function