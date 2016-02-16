'[path=\Projects\Project E\General Scripts]
'[group=General Scripts]
!INC Local Scripts.EAConstants-VBScript
!INC General Scripts.Util

' Script Name: LinkToCRMain
' Author: Geert Bellekens
' Purpose: Link Elemnents to a change
' Date: 2015-10-30
'
'


function linkItemToCR(selectedItem, selectedItems)
	dim groupProcessing
	groupProcessing = false
	'if the collection is given then we initialize the first item.
	if selectedItem is nothing then
		if not selectedItems is nothing then
			if selectedItems.Count > 0 then
				set selectedItem = selectedItems(0)
				if selectedItems.Count > 1 then
					groupProcessing = true
				end if
			end if
		end if
	end if
	if selectedItem is nothing then
		set selectedItem = Repository.GetContextObject()
	end if
	
	'get the select context item type
	dim selectedItemType
	selectedItemType = selectedItem.ObjectType
	select case selectedItemType
	case otElement, otPackage, otAttribute, otMethod, otConnector :
		'if the selectedItem is a package then we use the Element part of the package
		if selectedItemType = otPackage then
			dim selectedPackage
			'remember the package object
			set selectedPackage = selectedItem
			'get the element part of the package
			set selectedItem = selectedItem.Element
			'ask user if he wants to add the change to all owned elements of this package and its sub packages
			dim responsePackageTree
			responsePackageTree = Msgbox("Link all elements in package tree?", vbYesNoCancel+vbQuestion, "Link package tree?")
			'check the response
			select case responsePackageTree
				case vbYes
					'get all elements from package tree
					set selectedItems = getAllElementsInPackageTree(selectedPackage)
					'add the package itself as well
					selectedItems.Add selectedItem
					if selectedItems.Count > 1 then
						groupProcessing = true
					end if
				case vbCancel
					'user cancelled, stop altogether
					Exit function
			end select
		end if
		'get the logged in user
		Dim userLogin
		userLogin = getUserLogin
		dim lastCR as EA.Element
		set lastCR = nothing
		dim CRtoUse as EA.Element
		set CRtoUse = nothing
		set lastCR = getLastUsedCR(userLogin)
		'get most recent used CR by this user

		if not selectedItem is nothing then
			dim lastComments
			lastComments = vbNullString
			'if there is a last CR then we ask the user if we need to use that one
			if not lastCR is nothing then
				dim response
				if groupProcessing then
					response = Msgbox("Link all " & selectedItems.Count & " elements to change: """ & lastCR.Name & """?", vbYesNoCancel+vbQuestion, "Link to CR")
				elseif not isCRLinked(selectedItem,lastCR) then
					response = Msgbox("Link element """ & selectedItem.Name & """ to change: """ & lastCR.Name & """?", vbYesNoCancel+vbQuestion, "Link to CR")
				end if
				'check the response
				select case response
					case vbYes
						set CRToUse = lastCR
					case vbCancel
						'user cancelled, stop altogether
						Exit function
				end select
			end if
			'If there was no last CR, or the user didn't want to link that one we let the user choose one
			if CRToUse is nothing then
				dim CR_id 		
				CR_ID = Repository.InvokeConstructPicker("IncludedTypes=Change") 
				if CR_ID > 0 then
					set CRToUse = Repository.GetElementByID(CR_ID)
				end if
			else
				'user selected same change as last time. So he might want to reuse his comments as well
				lastComments = getLastUsedComment(userLogin)
			end if
			'if the CRtoUse is now selected then we link it to the selected element
			if not CRToUse is nothing then
				dim linkCounter
				linkCounter = 0
				'first check if this CR is not already linked
				if isCRLinked(selectedItem,CRToUse) and not groupProcessing then
					MsgBox "The CR was already linked to this item", vbOKOnly + vbExclamation ,"Already Linked" 
				else
					'get the comments to use
					dim comments
					comments = InputBox("Please enter comments for this change", "Change Comments",lastComments)
					if len(comments) > 2 then
						if groupProcessing then
							for each selectedItem in selectedItems
								'check the object type
								selectedItemType = selectedItem.ObjectType
								select case selectedItemType
								case otElement, otPackage, otAttribute, otMethod, otConnector :
									if not isCRLinked(selectedItem,CRToUse) then
										linkToCR selectedItem, selectedItemType, CRToUse, userLogin, comments
										linkCounter = linkCounter + 1
									end if
								end select
							next
							if linkCounter > 0 then
								MsgBox "Successfully linked " & selectedItems.Count & " elements to change """ & CRToUse.Name& """"  , vbOKOnly + vbInformation ,"Elements linked" 
							else
								MsgBox "No links created to change " & CRToUse.Name & "." & vbNewLine & "They are probably already linked" , vbOKOnly + vbExclamation ,"Already Linked" 
							end if
						else
							linkToCR selectedItem, selectedItemType, CRToUse, userLogin, comments
						end if
					else
						MsgBox "The CR has not been linked because no comment was provided", vbOKOnly + vbExclamation ,"No CR link" 
					end if
				end if
			end if
		end if
	case else
		MsgBox "Cannot link this type of element to a CR" & vbNewline & "Supported element types are: Element, Package, Attribute, Operation and Relation"
	end select
end function



function isCRLinked(item, CR)
	dim taggedValue as EA.TaggedValue
	isCRLinked = false
	for each taggedValue in item.TaggedValues
		if taggedValue.Value = CR.ElementGUID then
			isCRLinked = true
			exit for
		end if
	next
end function

function linkToCR(selectedItem, selectedItemType, CRToUse, userLogin, comments)
	Session.Output "CRToUse: " & CRToUse.Name & " userLogin: " & userLogin & " comments: " & comments
	dim crTag 
	set crTag = nothing
	set crTag = selectedItem.TaggedValues.AddNew("CR","")
	if not crTag is nothing then
		crTag.Value = CRToUse.ElementGUID
		crTag.Notes = "user=" & userLogin & ";" & _
					 "date=" & Year(Date) & "-" & Month(Date) & "-" & Day(Date) & ";" & _
					 "comments=" & comments
		crTag.Update
	end if
end function

function getLastUsedCR(userLogin)
	dim wildcard
	dim sqlDateString
	dim top
	dim limit
	if Repository.RepositoryType = "JET" then
		wildcard = "*"
		sqlDateString = " mid(tv.Notes, instr(tv.[Notes],'date=') + len('date='),10) "
		top = "top 1"
		limit = ""
	Elseif Repository.RepositoryType = "MYSQL" then
		wildcard = "%"
		sqlDateString = " substring(tv.Notes, instr('date=',tv.[Notes]) + length('date='),10) "
		top = ""
		limit = "limit 1"
	else 'SQL Server
		wildcard = "%"
		sqlDateString = " substring(tv.Notes, charindex('date=',tv.[Notes]) + len('date='),10) "
		top = "top 1"
		limit = ""
	end if
	dim sqlGetString
	sqlGetString = "select "& top &" o.Object_id " & _
					" from (t_objectproperties tv " & _
					" inner join t_object o on o.ea_guid = tv.VALUE) " & _
					" where tv.[Notes] like 'user=" & userLogin & ";" & wildcard & "' " & _
					" order by  " & sqlDateString & " desc, tv.PropertyID desc " & limit
	Session.Output "SQLGetString = " & sqlGetString
	dim CRs
	dim CR as EA.Element
	set CR = nothing
	'get the last CR
	set CRs = getElementsFromQuery(sqlGetString)
	if CRs.Count > 0 then
		set CR = CRs(0)
	end if

	set getLastUsedCR = CR
end function

function getLastUsedComment(userLogin)
	dim wildcard
	dim sqlDateString
	dim sqlCommentsString
	dim top
	dim limit
	if Repository.RepositoryType = "JET" then
		wildcard = "*"
		sqlDateString = " mid(tv.Notes, instr(tv.[Notes],'date=') + len('date='),10) "
		sqlCommentsString = " mid(tv.Notes, instr(tv.[Notes],'comments=') + len('comments=')) "
		top = "top 1"
		limit = ""
	Elseif Repository.RepositoryType = "MYSQL" then
		wildcard = "%"
		sqlDateString = " substring(tv.Notes, instr('date=',tv.[Notes]) + length('date='),10) "
		sqlCommentsString = " substring(tv.Notes, instr('comments=',tv.[Notes]) + length('comments='))  "
		top = ""
		limit = "limit 1"			
	Else 'SQL Server
		wildcard = "%"
		sqlDateString = " substring(tv.Notes, charindex('date=',tv.[Notes]) + len('date='),10) "
		sqlCommentsString = " substring(tv.Notes, charindex('comments=',tv.[Notes]) + len('comments='), datalength(tv.Notes))  "
		top = "top 1"
		limit = ""
	end if
	dim sqlGetString
	sqlGetString = "select " & top & sqlCommentsString & " as comments " & _
					" from (t_objectproperties tv " & _
					" inner join t_object o on o.ea_guid = tv.VALUE) " & _
					" where tv.[Notes] like 'user=" & userLogin & ";" & wildcard & "' " & _
					" order by  " & sqlDateString & " desc, tv.PropertyID desc " & limit
	dim queryResult 
	queryResult = Repository.SQLQuery(sqlGetString)
	Session.Output queryResult
	dim results
	results = convertQueryResultToArray(queryResult)
	if Ubound(results) > 0 then
		getLastUsedComment = results(0,0)
	else
		getLastUsedComment = vbNullString
	end if
end function