'[path=\Framework\OCL]
'[group=OCL]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
const outPutName = "Create Schema From Package"	

'create output tab
Repository.CreateOutputTab outPutName
Repository.ClearOutput outPutName
Repository.EnsureOutputVisible outPutName

sub main
	dim constraint
'	constraint = "LabelsCode::BF1   " & vbNewLine & _
'	"or self.Payload.Invoice_Type.content=LabelsCode::BF2   " & vbNewLine & _
'	"or self.Payload.Invoice_Type.content=LabelsCode::BK4	"
	constraint = "inv: self.Payload.Observation_Interval.Observation_Detail->forAll (Quantity_Quality.content=QuantityQualityCode::46 " & vbNewLine & _
    "or Quantity_Quality.content=QuantityQualityCode::81 " & vbNewLine & _
    "or Quantity_Quality.content=QuantityQualityCode::56 " & vbNewLine & _
    "or  Quantity_Quality.content=QuantityQualityCode::86" & vbNewLine & _
    "or  Quantity_Quality.content=QuantityQualityCode::125)"



	'first remove the comments
	dim trimmedConstraint
	Dim regExp		
	Set regExp = CreateObject("VBScript.RegExp")
	regExp.Global = True   
	regExp.IgnoreCase = False
	regExp.Pattern = "(--.*)"
	trimmedConstraint = regExp.Replace(constraint, "")
	'then group by individual OCL statement
	dim statements
	statements = split(trimmedConstraint, "inv:")
	dim statement
	dim i
	i = 0
	for each statement in statements
		Repository.WriteOutput outPutName, now() &  " processing OCL statement " & statement, 0 
		dim newOCL
		dim matches
		'remove leading whitespace
		regExp.Pattern = "^\s*"
		statement = regExp.Replace(statement, "")
		'remove trailing whitespace
		regExp.Pattern = "\s*$"
		statement = regExp.Replace(statement, "")
		'group into identifier (1), operator(2) and value(3)
		'regExp.Pattern = "(^[\s\w\.]+(?=(->size\(\)=?|->notEmpty\(\)|->forAll ?\(|=)(?=([\s\S]+$))))"
		'regExp.Pattern ="(^[\s\w\.]+(?=->size\(\)=?|->notEmpty\(\)|->forAll ?\(|=(?=[\s\S]+$))(->size\(\)=?|->notEmpty\(\)|->forAll ?\(|=)(?=[\s\S]+$))([\s\S]+$)"
		'regExp.Pattern = "(^.*?(?=->size\(\)=?|->notEmpty\(\)|->forAll ?\(|[^)]=(?=[\s\S]+$)))(->size\(\)=?|->notEmpty\(\)|->forAll ?\(|[^)]=)(?=[\s\S]+$)([\s\S]+$)"
		regExp.Pattern = "(^.*?)(->size\(\)=?|->notEmpty\(\)|->forAll ?\(|=)(?=[\s\S]+$)?([\s\S]+$)?"
		'regExp.Pattern = "(^[\s\w\.]+?(?==)(=))"
		set matches = regExp.Execute(statement)
		dim match
		for each match in matches
			dim subMatch
			dim j
			for j = 0 To match.SubMatches.Count-1
			'for each subMatch in match.SubMatches
				'debug
				Repository.WriteOutput outPutName, now() &  " submatch found: " & match.SubMatches(j) , 0 
			next
		next
	next
end sub

main