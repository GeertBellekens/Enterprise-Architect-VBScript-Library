
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