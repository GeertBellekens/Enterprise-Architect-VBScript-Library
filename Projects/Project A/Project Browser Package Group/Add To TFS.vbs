'[path=\Projects\Project A\Project Browser Package Group]
'[group=Project Browser Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Add to TFS
' Author: Geert Bellekens
' Purpose: Adds the whole package tree to TFS
' Date: 2017-10-19
'

const versionControlID = "TFS_EA_MODEL"
const outPutName = "Add to TFS"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get the selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	'let the user know we started
	Repository.WriteOutput outPutName, now() & " Starting adding to TFS for package '"& package.Name &"'", 0
	'ask the user if he is sure
	dim userIsSure
	userIsSure = Msgbox("Do you really want to add package '" &package.Name & "' to TFS?", vbYesNo+vbQuestion, "Add package to TFS?")
	if userIsSure = vbYes then
		on error resume next
		'get the subfolder from the user
		dim subFolder
		subFolder = getVersionControlSubFolderPath()
		'check if error was raised
		if Err.number <> 0 then
			Err.clear
			Repository.WriteOutput outPutName, now() & " Cancelled by user", 0
			exit sub
		end if
		on error goto 0
		'add the selected package to version control
		addToVersionControl package, subfolder
	end if
	'let the user know it is finished
	Repository.WriteOutput outPutName, now() & " Finished adding to TFS for package '"& package.Name &"'", 0
end sub

function addToVersionControl(package, subfolder)
	'first process subPackages
	dim subPackage
	for each subPackage in package.Packages
		addToVersionControl subPackage, subfolder
	next
	
	if not package.IsVersionControlled then
		Repository.WriteOutput outPutName, now() & " Adding package '"& package.Name &"'", 0
		'then add this package to version control
		package.VersionControlAdd versionControlID, _
						subfolder & package.PackageGUID + cleanFileName(package.Name) & ".xml", _
						"Initial addition of package " & package.Name , _
						false
	else
		'tell the user this package is already added to version control.
		Repository.WriteOutput outPutName, now() & " Error: Can't add package '"& package.Name &"'", 0
	end if
end function

function getVersionControlSubFolderPath()
	'get the base folder for this version control configuration
	dim baseFolderName
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
		dim currentVCID
		currentVCID = getValueForkey(pathLine, "id")
		if currentVCID = versionControlID then
			baseFolderName = getValueForkey(pathLine, "path")
			exit for
		end if
	next
	'let the user select a subFolder
	dim userSelectedFolder
    set userSelectedFolder = new FileSystemFolder
	set userSelectedFolder = userSelectedFolder.getUserSelectedFolder(baseFolderName)
	if userSelectedFolder is nothing then
		Err.Raise vbObjectError + 10, "Add to TFS", "Folder Selection Cancelled"
		exit function
	end if
	'return the folder
	getVersionControlSubFolderPath = mid(userSelectedFolder.FullPath, len(baseFolderName) + 1 )
	'remove leading "\"
	if len (getVersionControlSubFolderPath) > 0 then
		if left(getVersionControlSubFolderPath,1) = "\" then
			getVersionControlSubFolderPath = mid(getVersionControlSubFolderPath,2) & "\"
		end if
	end if
end function


main