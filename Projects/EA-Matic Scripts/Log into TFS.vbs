'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Log into TFS
' Author: Geert Bellekens
' Purpose: Executes a "tf get" command at file open to make sure the connection to TFS is refreshed if needed.
' Date: 2019-11-15
'
'EA-Matic

function EA_FileOpen()
	Dim objShell
	Set objShell = CreateObject ("WScript.shell")
	objShell.run "cmd /c CD ""C:\Program Files (x86)\Microsoft Visual Studio\2019\TeamExplorer\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\"" & tf get"
	Set objShell = Nothing
end function