'[path=\Projects\Project H\Hampden Scripts]
'[group=Hampden Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include


'
' Script Name: Export ValidationRules
' Author: Geert Bellekens
' Purpose: Export the ValidationRules to a csv file to be imported by ODS
' Date: 2020-09-22
'
const outputName = "Export ValidationRules"
const rulesPackageGUID = "{190B02C3-7A82-4df5-90DB-171243B3818F}"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'inform user
	Repository.WriteOutput outPutName, now() & " Starting Export Validation Rules", 0
	'start the actual work
	exportValidationRules
	'inform user
	Repository.WriteOutput outPutName, now() & " Finished Export Validation Rules", 0
end sub

function exportValidationRules
	Repository.WriteOutput outPutName, now() & " Getting data", 0
	dim rulesPackage as EA.Package
	set rulesPackage = Repository.GetPackageByGuid(rulesPackageGUID)
	dim rulesPackageTreeIDs
	rulesPackageTreeIDs = getPackageTreeIDString(rulesPackage)
	'get data in memory
	dim sqlGetData
	sqlGetData = 	"select o.name, o.Note, tvr.value as RuleID, tvb.Value as BIVP        " & vbNewLine & _
					" , tvc.Value as CorrectiveAction, tvp.Value as SanityPriority        " & vbNewLine & _
					" , tvs.Value as SuggestedAction                                      " & vbNewLine & _
					" from t_object o                                                     " & vbNewLine & _
					" left join t_objectproperties tvr on tvr.Object_ID = o.Object_ID     " & vbNewLine & _
					" 								and tvr.Property = 'RuleID'           " & vbNewLine & _
					" left join t_objectproperties tvb on tvb.Object_ID = o.Object_ID     " & vbNewLine & _
					" 								and tvb.Property = 'BIVP'             " & vbNewLine & _
					" left join t_objectproperties tvc on tvc.Object_ID = o.Object_ID     " & vbNewLine & _
					" 								and tvc.Property = 'CorrectiveAction' " & vbNewLine & _
					" left join t_objectproperties tvp on tvp.Object_ID = o.Object_ID     " & vbNewLine & _
					" 								and tvp.Property = 'SanityPriority'   " & vbNewLine & _
					" left join t_objectproperties tvs on tvs.Object_ID = o.Object_ID     " & vbNewLine & _
					" 								and tvs.Property = 'SuggestedAction'  " & vbNewLine & _
					" where o.Stereotype = 'DQM_ValidationRule'                           " & vbNewLine & _
					" and o.Package_ID in (" & rulesPackageTreeIDs & ")                   " & vbNewLine & _
					" order by cast(isnull(tvr.Value,999) as int), o.name                 "
					
	dim data
	set data = getArrayListFromQuery(sqlGetData)
	'write data to file
	Repository.WriteOutput outPutName, now() & " Exporting data to file", 0
	dim csvFile
	set csvFile = new CSVFile
	csvFile.Contents = data
	'ask user if we need this for production
	dim response
		response = msgbox("Export for production?" , vbYesNo+vbQuestion, "Production or Test")
		if response = vbYes then
			csvFile.FullPath = "\\Omnibus\PRD_EA\ValidationRulesOutput.csv"
			csvFile.Save
			csvFile.FullPath = "\\Omnibus\GA_Evaluation_EA\ValidationRulesOutput.csv"
			csvFile.Save
		else
			'let the user select a file for himself
			dim selectedFolder
			set selectedFolder = new FileSystemFolder
			set selectedFolder = selectedFolder.getUserSelectedFolder("")
			csvFile.FullPath = selectedFolder.FullPath & "\" & "ValidationRulesOutput.csv"
			csvFile.Save
		end if

end function

main