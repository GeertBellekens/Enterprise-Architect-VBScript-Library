'[path=\Framework\Utils]
'[group=Utils]
'Author: Geert Bellekens
'Date: 2015-12-07

'returns the SQL wildcard depending on the type of repository
function getWC()
	if Repository.RepositoryType = "JET" then
		getWC = "*"
	else
		getWC = "%"
	end if
end function

'escapes a literal string so it can be inserted using sql
function escapeSQLString(inputString)
	'replace the single quotes with two single quotes for all db types
	escapeSQLString = replace(inputString, "'","''")
	'dbspecifics
	select case Repository.RepositoryType
		case "POSTGRES"
			' replace backslash "\" by double backslash "\\"
			escapeSQLString = replace(escapeSQLString,"\","\\")
		case "JET"
			'replace pipe character | by '& chr(124) &'
			escapeSQLString = replace(escapeSQLString,"|", "'& chr(124) &'")
	end select
end function
