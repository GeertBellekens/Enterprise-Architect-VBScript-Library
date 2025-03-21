'[path=\Projects\Project B\Package Group]
'[group=Package Group]
option explicit


!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Cleanup Split Labels
' Author: Geert Bellekens
' Purpose: Remove the split labels from the notes of elements, attributes, connectors and 
' Date: 2018-09-24
'
const outPutName = "Cleanup Split Labels"

function Main ()
	'get selected package
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	if not selectedPackage is Nothing then
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'set timestamp
		Repository.WriteOutput outPutName, now() & " Starting Cleanup split labels for package '" & selectedPackage.Name & "'" , selectedPackage.Element.ElementID
		'execute the cleanup
		cleanupSplitLabels selectedPackage
		'set timestamp
		Repository.WriteOutput outPutName, now() & " Finished Cleanup split labels for package '" & selectedPackage.Name & "'" , selectedPackage.Element.ElementID	
	end if
	'reload modelview
	Repository.RefreshModelView 0
end function

function cleanupSplitLabels(selectedPackage)
	'get the packageTreeIDs
	dim packageTreeIDs
	packageTreeIDs = getPackageTreeIDString(selectedPackage)
	dim sqlUpdate
	
	'-- 1. XSDcomplexTypes
	Repository.WriteOutput outPutName, now() & " Processing 1. XSDcomplexTypes" , selectedPackage.Element.ElementID	
	sqlUpdate = "UPDATE t_object                                                                        " &  VBNewLine & _
				"    SET Note =                                                                         " &  VBNewLine & _
				" 	   CASE SUBSTRING(Note, 1, 3)                                                       " &  VBNewLine & _
				" 		  WHEN 'all' THEN SUBSTRING(Note,  6, LEN(Note)) --all                          " &  VBNewLine & _
				" 		  WHEN 'lif' THEN SUBSTRING(Note,  7, LEN(Note)) --life                         " &  VBNewLine & _
				" 		  WHEN 'non' THEN SUBSTRING(Note, 10, LEN(Note)) --nonlife                      " &  VBNewLine & _
				" 		  WHEN 'nlf' THEN SUBSTRING(Note,  9, LEN(Note)) --nlfire                       " &  VBNewLine & _
				" 		  WHEN 'nlv' THEN SUBSTRING(Note, 10, LEN(Note)) --nlvaria, nlvehic, nlvehfi    " &  VBNewLine & _
				" 		  ELSE Note                                                                     " &  VBNewLine & _
				" 		END                                                                             " &  VBNewLine & _
				"   from t_object       o                                                               " &  VBNewLine & _
				"  where o.Stereotype = 'XSDcomplexType'                                                " &  VBNewLine & _
				"    and (  o.Note like 'all%'                                                          " &  VBNewLine & _
				"        OR o.Note like 'lif%'                                                          " &  VBNewLine & _
				" 	   OR o.Note like 'non%'                                                            " &  VBNewLine & _
				" 	   OR o.Note like 'nlf%'                                                            " &  VBNewLine & _
				" 	   OR o.Note like 'nlv%')                                                           " &  VBNewLine & _
				" 	 and  o.Package_ID in ("& packageTreeIDs &")" 
	Repository.Execute  sqlUpdate
	
	'-- 2. XSDelement
	Repository.WriteOutput outPutName, now() & " Processing 2. XSDelement" , selectedPackage.Element.ElementID
	sqlUpdate = "UPDATE t_attribute                                                            			   " &  VBNewLine & _
				"    SET Notes =                                                                           " &  VBNewLine & _
				"        CASE SUBSTRING(a.Notes, 1, 3)                                                     " &  VBNewLine & _
				" 		 WHEN 'all' THEN SUBSTRING(a.Notes,  6, LEN(a.Notes)) --all                        " &  VBNewLine & _
				" 		 WHEN 'lif' THEN SUBSTRING(a.Notes,  7, LEN(a.Notes)) --life                       " &  VBNewLine & _
				" 		 WHEN 'non' THEN SUBSTRING(a.Notes, 10, LEN(a.Notes)) --nonlife                    " &  VBNewLine & _
				" 		 WHEN 'nlf' THEN SUBSTRING(a.Notes,  9, LEN(a.Notes)) --nlfire                     " &  VBNewLine & _
				" 		 WHEN 'nlv' THEN SUBSTRING(a.Notes, 10, LEN(a.Notes)) --nlvaria, nlvehic, nlvehfi  " &  VBNewLine & _
				" 		 ELSE a.Notes                                                                      " &  VBNewLine & _
				" 	   END                                                                                 " &  VBNewLine & _
				"   FROM t_attribute    a                                                                  " &  VBNewLine & _
				"   INNER JOIN t_object       o on o.Object_ID  = a.Object_ID                              " &  VBNewLine & _
				"  WHERE o.Stereotype = 'XSDcomplexType'                                                   " &  VBNewLine & _
				"    AND (  a.Notes like 'all%'                                                            " &  VBNewLine & _
				"        OR a.Notes like 'lif%'                                                            " &  VBNewLine & _
				"        OR a.Notes like 'non%'                                                            " &  VBNewLine & _
				"        OR a.Notes like 'nlf%'                                                            " &  VBNewLine & _
				"        OR a.Notes like 'nlv%' )                                                          " &  VBNewLine & _
				" 	 and  o.Package_ID in ("& packageTreeIDs &")"
	Repository.Execute  sqlUpdate
    
	' -- 3. XSD Simple types
	Repository.WriteOutput outPutName, now() & " Processing 3. XSD Simple types" , selectedPackage.Element.ElementID
	sqlUpdate = "UPDATE t_object                                                                                  " &  VBNewLine & _
				"    SET Note = SUBSTRING(Note, CHARINDEX(')', Note, 1)+3, LEN(Note) - CHARINDEX(')', Note, 1)+3) " &  VBNewLine & _
				"   FROM t_object       o                                                                         " &  VBNewLine & _
				"  WHERE o.Stereotype = 'XSDsimpleType'                                                           " &  VBNewLine & _
				"    AND (  o.Note like '(all%'                                                                   " &  VBNewLine & _
				"        OR o.Note like '(lif%'                                                                   " &  VBNewLine & _
				" 	   OR o.Note like '(non%'                                                                     " &  VBNewLine & _
				" 	   OR o.Note like '(nlf%'                                                                     " &  VBNewLine & _
				" 	   OR o.Note like '(nlv%')                                                                    " &  VBNewLine & _
				" 	 and  o.Package_ID in ("& packageTreeIDs &")"
	Repository.Execute  sqlUpdate
	
	' -- 4. Enumerations
	Repository.WriteOutput outPutName, now() & " Processing 4. Enumerations" , selectedPackage.Element.ElementID
	sqlUpdate = " UPDATE t_object                                                                                 " &  VBNewLine & _
				"    SET Note = SUBSTRING(Note, CHARINDEX(')', Note, 1)+3, LEN(Note) - CHARINDEX(')', Note, 1)+3) " &  VBNewLine & _
				"   FROM t_object       o                                                                         " &  VBNewLine & _
				"  WHERE o.Object_Type = 'Enumeration'                                                            " &  VBNewLine & _
				"    AND (  o.Note like '(all%'                                                                   " &  VBNewLine & _
				"        OR o.Note like '(lif%'                                                                   " &  VBNewLine & _
				" 	   OR o.Note like '(non%'                                                                     " &  VBNewLine & _
				" 	   OR o.Note like '(nlf%'                                                                     " &  VBNewLine & _
				" 	   OR o.Note like '(nlv%')                                                                    " &  VBNewLine & _
				" 	 and  o.Package_ID in ("& packageTreeIDs &")"
	Repository.Execute  sqlUpdate

	' -- 5. Enumeration values
	Repository.WriteOutput outPutName, now() & " Processing 5. Enumeration values" , selectedPackage.Element.ElementID
	sqlUpdate = "UPDATE t_attribute                                                                                                                                                                                                         " &  VBNewLine & _
				"    SET Notes = SUBSTRING(a.Notes, 3, CHARINDEX('-', a.Notes)-4) + ' - ' + SUBSTRING(a.Notes, CHARINDEX('-', a.Notes, CHARINDEX('-', a.Notes)+1)+2, LEN(a.Notes) - CHARINDEX('-', a.Notes, CHARINDEX('-', a.Notes)+1)-3)   " &  VBNewLine & _
				"   FROM t_object       o                                                                                                                                                                                                   " &  VBNewLine & _
				"   JOIN t_attribute    a on a.Object_ID  = o.Object_ID                                                                                                                                                                     " &  VBNewLine & _
				"  WHERE o.Object_Type = 'Enumeration'                                                                                                                                                                                      " &  VBNewLine & _
				"    AND a.Notes like '[[]%'                                                                                                                                                                                                " &  VBNewLine & _
				" 	 and  o.Package_ID in ("& packageTreeIDs &")"
	Repository.Execute  sqlUpdate
	
	' -- 6. Associations
	Repository.WriteOutput outPutName, now() & " Processing 6. Associations" , selectedPackage.Element.ElementID
	sqlUpdate = " UPDATE t_connector                                                                                " &  VBNewLine & _
				"    SET DestRoleNote =                                                                             " &  VBNewLine & _
				"        CASE SUBSTRING(DestRoleNote, 1, 3)                                                         " &  VBNewLine & _
				" 		 WHEN 'all' THEN SUBSTRING(DestRoleNote,  6, LEN(DestRoleNote)) --all                       " &  VBNewLine & _
				" 		 WHEN 'lif' THEN SUBSTRING(DestRoleNote,  7, LEN(DestRoleNote)) --life                      " &  VBNewLine & _
				" 		 WHEN 'non' THEN SUBSTRING(DestRoleNote, 10, LEN(DestRoleNote)) --nonlife                   " &  VBNewLine & _
				" 		 WHEN 'nlf' THEN SUBSTRING(DestRoleNote,  9, LEN(DestRoleNote)) --nlfire                    " &  VBNewLine & _
				" 		 WHEN 'nlv' THEN SUBSTRING(DestRoleNote, 10, LEN(DestRoleNote)) --nlvaria, nlvehic, nlvehfi " &  VBNewLine & _
				" 		 ELSE DestRoleNote                                                                          " &  VBNewLine & _
				"        END                                                                                        " &  VBNewLine & _
				"   FROM t_connector    a                                                                           " &  VBNewLine & _
				"   JOIN t_object       s on s.Object_ID  = a.Start_Object_ID                                       " &  VBNewLine & _
				"   JOIN t_object       d on d.Object_ID  = a.End_Object_ID                                         " &  VBNewLine & _
				"  WHERE a.DestRoleNote is not null                                                                 " &  VBNewLine & _
				"    AND (  a.DestRoleNote like 'all%'                                                              " &  VBNewLine & _
				"        OR a.DestRoleNote like 'lif%'                                                              " &  VBNewLine & _
				" 	   OR a.DestRoleNote like 'non%'                                                                " &  VBNewLine & _
				" 	   OR a.DestRoleNote like 'nlf%'                                                                " &  VBNewLine & _
				" 	   OR a.DestRoleNote like 'nlv%')                                                               " &  VBNewLine & _
				" 	 and  s.Package_ID in ("& packageTreeIDs &")"
	Repository.Execute  sqlUpdate	
end function

'get the package id string of the given package tree
function getPackageTreeIDString(package)
	'initialize at "0"
	getPackageTreeIDString = "0"
	dim packageTree
	dim currentPackage as EA.Package
	if not package is nothing then
		'get the whole tree of the selected package
		set packageTree = getPackageTree(package)
		' get the id string of the tree
		getPackageTreeIDString = makePackageIDString(packageTree)
	end if 
end function

'returns an ArrayList of the given package and all its subpackages recursively
function getPackageTree(package)
	dim packageList
	set packageList = CreateObject("System.Collections.ArrayList")
	addPackagesToList package, packageList
	set getPackageTree = packageList
end function

'add the given package and all subPackges to the list (recursively
function addPackagesToList(package, packageList)
	dim subPackage as EA.Package
	'add the package itself
	packageList.Add package
	'add subpackages
	for each subPackage in package.Packages
		addPackagesToList subPackage, packageList
	next
end function

'make an id string out of the package ID of the given packages
function makePackageIDString(packages)
	dim package as EA.Package
	dim idString
	idString = ""
	dim addComma 
	addComma = false
	for each package in packages
		if addComma then
			idString = idString & ","
		else
			addComma = true
		end if
		idString = idString & package.PackageID
	next 
	'if there are no packages then we return "0"
	if idString = "" then
		idString = "0"
	end if
	'return idString
	makePackageIDString = idString
end function

main