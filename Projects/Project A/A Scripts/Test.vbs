'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

sub main
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	dim startTime
	dim endTime
	
	startTime = Timer()
	dim slowPackageTreeIDs
	slowPackageTreeIDs = getPackageTreeIDString(package)
	endTime = Timer()
	Session.Output "Slow getPackageTreeIDString: " & FormatNumber(EndTime - StartTime, 3)
	
	startTime = Timer()
	dim fastPackageTreeIDs
	fastPackageTreeIDs = getPackageTreeIDStringFast(package)
	endTime = Timer()
	Session.Output "Fast getPackageTreeIDString: " & FormatNumber(EndTime - StartTime, 3)
	
	Session.Output "Slow package IDs: " & slowPackageTreeIDs
	Session.Output "Fast package IDs: " & fastPackageTreeIDs
end sub


function getPackageTreeIDStringFast(package)
	dim allPackageTreeIDs 
	set allPackageTreeIDs = CreateObject("System.Collections.ArrayList")
	dim parentPackageIDs
	set parentPackageIDs = CreateObject("System.Collections.ArrayList")
	parentPackageIDs.Add package.PackageID
	getPackageTreeIDsFast allPackageTreeIDs, parentPackageIDs
	'return
	getPackageTreeIDStringFast = Join(allPackageTreeIDs.ToArray,",")
end function

function getPackageTreeIDsFast(allPackageTreeIDs, parentPackageIDs)
	if parentPackageIDs.Count = 0 then
		if allPackageTreeIDs.Count = 0 then
			'make sure there is at least a 0 in the allPackageTreeIDs
			allPackageTreeIDs.Add "0"
		end if
		'then exit
		exit function
	end if
	'add the parent package ids
	allPackageTreeIDs.AddRange(parentPackageIDs)
	'get the child package IDs
	dim sqlGetPackageIDs
	sqlGetPackageIDs = "select p.Package_ID from t_package p where p.Parent_ID in (" & Join(parentPackageIDs.ToArray, ",") & ")"
	dim queryResult
	set queryResult = getVerticalArrayListFromQuery(sqlGetPackageIDs)
	if queryResult.Count > 0 then
		dim childPackageIDs
		set childPackageIDs = queryResult(0)
		'call recursive function with child package id's
		getPackageTreeIDsFast allPackageTreeIDs, childPackageIDs
	end if
end function



main