'[path=\Framework\Utils]
'[group=Utils]
' Script Name: ScriptLogger
' Author: Geert Bellekens
' Purpose: used to log the execution of scripts to keep track of which scripts are being executed by whom
' Date: 2019-10-16
!INC Wrappers.Include

const logfilePath = "G:\Projects\80 Enterprise Architect\Output\ScriptLogs"

function LogScriptExecution (scriptName, info)
	'get the username
	dim userName
	userName = replace(Repository.GetCurrentLoginUser(false), "\", "_")
	'scriptlogs are saved in a file per user, in a folder per year
	dim fileName
	fileName = logfilePath & "\" & Year(now) & "\" & userName & "_scriptlog.log"
	'append to the file
	dim logfile
	set logFile = new TextFile
	logFile.fullPath = fileName
	logFile.append now() & ";" & userName & ";" & Repository.ConnectionString & ";" _
					&  scriptName & ";" & info & vbNewLine
	'msgbox logfile.fullPath
end function

function test()
	LogScriptExecution "LogScriptExecution", "info"
end function

test