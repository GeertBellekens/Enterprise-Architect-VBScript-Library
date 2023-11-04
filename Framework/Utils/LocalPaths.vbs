'[path=\Framework\utils]
'[group=Utils]

!INC Utils.Util

'
' Given an id "byId" return that local path defined in %APPDATA%\Sparx Systems\EA\paths.txt
' if no id matches then return an empty string
'
function LocalPathsToPathForId(byId)
    LocalPathsToPathForId = ""

    dim shell
    Set shell = CreateObject( "WScript.Shell" )
    dim appDataPath
    appDataPath = shell.ExpandEnvironmentStrings("%APPDATA%")
    dim pathsFile
    set pathsFile = new TextFile
    pathsFile.FullPath = appDataPath & "\Sparx Systems\EA\paths.txt"
    pathsFile.loadContents
    dim pathsLines
    pathsLines = Split(pathsFile.Contents, vbCrLf)
    dim pathLine
    for each pathLine in pathsLines
        dim pathId
        pathId = getValueForkey(pathLine, "id")
        if pathId = byId then
            LocalPathsToPathForId = getValueForkey(pathLine, "path")
            exit function
        end if
    next
end function