'[path=\Framework\ModelValidation]
'[group=ModelValidation]

'
' Given a baseId and a localId, make a ruleId by concatenation
' MVR{baseId + localId}
'
function makeId(baseId, localId)
	makeId = "MVR" & (baseId + localId)
end function