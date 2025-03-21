'[path=\Framework\Tools\Script Management]
'[group=Script Management]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC EAScriptLib.VBScript-GUID

' Author: Geert Bellekens
' Purpose: Loads scripts from the file systems and stores them in Enterprise Architect
' Date: 2015-12-07
'
sub main
	dim selectedFolder,file, allScripts, allGroups,script, overwriteExisting
	set selectedFolder = new FileSystemFolder
	set selectedFolder = selectedFolder.getUserSelectedFolder("")
	overwriteExisting = "undecided"
	if not selectedFolder is nothing then
		set allGroups = Nothing
		set script = new Script
		'first get all existing scripts and groups
		set allScripts = Script.getAllScripts(allGroups)
		'get the scripts from the folder and its subfolders
		getScriptsFromFolder selectedFolder, allGroups, allScripts, overwriteExisting
	end if
end sub

'gets all the scripts from the given folder and its subfolders (if any)
function getScriptsFromFolder(selectedFolder, allGroups, allScripts, overwriteExisting)
	dim script, subFolder
	for each file in selectedFolder.TextFiles
		Session.Output "FileName: " & file.FileName
		'Session.Output "Code: " & file.Contents
		set script = getScriptFromFile(file, allGroups, allScripts,overwriteExisting)
		if overwriteExisting = vbCancel then
			exit for
		end if
	next
	'then process subfolders
	if not overwriteExisting = vbCancel then
		for each subFolder in selectedFolder.SubFolders
			getScriptsFromFolder subFolder, allGroups, allScripts, overwriteExisting
		next
	end if
end function

function getScriptFromFile(file, allGroups, allScripts,overwriteExisting)
	dim script, newScript, foundMatch, newScriptGroupName, group, foundGroup
	foundMatch = false
	foundGroup = false
	set group = nothing
	set script = Nothing
	if file.Extension = "vbs" then
		for each script in allScripts
			set newScript = new Script
			newScript.Name = file.FileNameWithoutExtension
			newScript.Code = file.Contents
			newScriptGroupName = newScript.GroupInNameCode 
			'if the groupname was not found in the code we use the name of the package
			if len(newScriptGroupName) = 0 then
				newScriptGroupName = file.Folder.Name
			end if
			'check the name of the script
			if script.Name = newScript.Name then
				'check if there is a groupname defined in the file
				if script.Group.Name = newScriptGroupName then
					'we have a match
					foundMatch = true
					set group = script.Group
					exit for
				end if
			end if
		next
		if not foundMatch then 
			'script did not exist yet
			'figure out if the group exists already
			for each group in allGroups.Items
				if group.Name = newScriptGroupName then
					'found the group
					'add the group to the new script
					newScript.Group = group
					checkGroupTypeAndOverwrite group, newScript
					foundGroup = true
					exit for
				end if
			next
			'if the group doesn't exist yet we have to create it
			if not foundGroup then
				set group = new ScriptGroup
				group.Name = newScriptGroupName
				group.GUID = GUIDGenerateGUID()
				group.GroupType = newScript.GroupType
				'create the Group in the database
				group.Create
				'refresh allGroups
				Session.Output "allGroups.Count before: " & allGroups.Count
				set allGroups = group.GetAllGroups()
				Session.Output "allGroups.Count after: " & allGroups.Count
				'add the group to the script
				newScript.Group = group
			end if
			'Now we have to create the script
			newScript.GUID = GUIDGenerateGUID()
			newScript.Create
			set script = newScript
		else
			if overwriteExisting = "undecided" then
				overwriteExisting = Msgbox("Do you want to update existing scripts?", vbYesNoCancel+vbQuestion, "Update existing scripts")
			end if
			if overwriteExisting = vbYes then
				checkGroupTypeAndOverwrite group, newScript
				script.Code = newScript.Code
				script.Update
			end if
		end if

	end if
	set getScriptFromFile = script
end function

sub checkGroupTypeAndOverwrite(group, newScript)
	if group.GroupType <> newScript.GroupType then
		dim updateGroupType
		updateGroupType = Msgbox("Update Group " & group.Name & " from GroupType=" & group.GroupType & " to new GroupType=" & newScript.GroupType & "?", vbYesNoCancel+vbQuestion, "Script group type does not match Group group type")
		if updateGroupType = vbYes then
			group.GroupType = newScript.GroupType
			group.Update
		end if
	end if
end sub

main