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
	'get the automatic or user selected version control ID
	dim versionControlID
	versionControlID = getUserSelectedVersionControlID(package)
	if not len(versionControlID) > 0 then
		Msgbox "No version control ID selected or present in the model"
		exit sub
	end if
	'ask the user if he is sure
	dim userIsSure
	userIsSure = Msgbox("Do you really want to add package '" &package.Name & "' to '" & versionControlID &  "' in TFS?", vbYesNo+vbQuestion, "Add package to TFS?")
	if userIsSure = vbYes then
		on error resume next
		'get the subfolder from the user
		dim subFolder
		subFolder = getVersionControlSubFolderPath(versionControlID)
		'check if error was raised
		if Err.number <> 0 then
			Err.clear
			Repository.WriteOutput outPutName, now() & " Cancelled by user", 0
			exit sub
		end if
		on error goto 0
		'add the selected package to version control
		addToVersionControl package, subfolder, versionControlID
	end if
	'let the user know it is finished
	Repository.WriteOutput outPutName, now() & " Finished adding to TFS for package '"& package.Name &"'", 0
end sub

function getUserSelectedVersionControlID(package)
	'check if the parent package has a version control ID
	dim versionControlIDs 
	set versionControlIDs = getVersionControlIDsForPackages(package.ParentID)
	if versionControlIDs.Count = 1 then
		'return this one
		getUserSelectedVersionControlID = versionControlIDs(0)
		exit function
	elseif versionControlIDs.Count = 0 then
		'get all version control ID's in this model
		set versionControlIDs = getVersionControlIDs()
	end if
	if versionControlIDs.Count <= 0 then
		'no version control ID's here
		exit function
	end if
	'build string
	dim selectMessage
	selectMessage = "Please enter the number of the configuration"
	dim i
	i = 0
	dim versionControlID
	for each versionControlID in versionControlIDs
		i = i + 1
		selectMessage = selectMessage & vbNewLine & i & ": " & versionControlID
	next
	dim response
	response = InputBox(selectMessage, "Select Version control ID", "1" )
	if isNumeric(response) then
		if Cstr(Cint(response)) = response then 'check if response is integer
			dim selectedID
			selectedID = Cint(response) - 1
			if selectedID >= 0 and selectedID < versionControlIDs.Count then
				'return the version control ID
				getUserSelectedVersionControlID = versionControlIDs(selectedID)
			end if
		end if
	end if
end function

function addToVersionControl(package, subfolder, versionControlID)
	'first process subPackages
	dim subPackage
	for each subPackage in package.Packages
		addToVersionControl subPackage, subfolder, versionControlID
	next
	
	if not package.IsVersionControlled _
	  and lcase(package.Element.Stereotype) = "xsdschema" _
	  and lcase(right(package.Name, 5)) <> "types" then
		Repository.WriteOutput outPutName, now() & " Adding package '"& package.Name &"'", 0
		'then add this package to version control
		package.VersionControlAdd versionControlID, _
						subfolder & cleanFileName(package.Name) & package.PackageGUID & ".xml", _
						"Initial addition of package " & package.Name , _
						false
	else
		'tell the user this package is already added to version control.
		Repository.WriteOutput outPutName, now() & " Skipped package '"& package.Name &"'", 0
	end if
end function

function getVersionControlSubFolderPath(versionControlID)
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