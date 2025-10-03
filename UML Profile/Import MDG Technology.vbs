'[group=UML Profile]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'


' ===== Insert a local file into SQL Server (no OPENROWSET) =====


'---- CONFIG ----
Dim server, database, useIntegratedSecurity, uid, pwd
server = "DESKTOP-BGN5EL4"           ' e.g. "MY-SQL01" or "MY-SQL01\SQLEXPRESS"
database = "TMF"
useIntegratedSecurity = True                      ' False to use SQL auth
uid = "sa"                                        ' if useIntegratedSecurity=False
pwd = "yourStrong(!)Password"                     ' if useIntegratedSecurity=False

Dim filePath, docName, elementId, elementType, docType
filePath   = "G:\My Drive\Klanten Bellekens IT\Atrias\test import MDG\TestRedefine MDG.zip"           ' <<-- your local file
docName    = "TRD"                                 ' nvarchar(100)
elementId  = "TECHNOLOGY"                          ' nvarchar(40)
elementType= "TECHNOLOGY"                          ' nvarchar(50)
docType    = "TECHNOLOGY"                                 ' nvarchar(100)
'--------------


'---- ADO constants (no adovbs.inc) ----
Const adCmdText       = 1
Const adParamInput    = 1
Const adVarWChar      = 202
Const adVarChar       = 200
Const adVarBinary     = 204
Const adLongVarBinary = 205
Const adTypeBinary    = 1
'---------------------------------------

function main
	' Build connection string (MSOLEDBSQL preferred; fallback to SQLOLEDB if needed)
	Dim connStr
	If useIntegratedSecurity Then
	  connStr = "Provider=MSOLEDBSQL;Server=" & server & ";Database=" & database & ";Trusted_Connection=Yes;"
	Else
	  connStr = "Provider=MSOLEDBSQL;Server=" & server & ";Database=" & database & ";User ID=" & uid & ";Password=" & pwd & ";"
	End If

	Dim cn: Set cn = CreateObject("ADODB.Connection")
	cn.Open connStr
'	If Err.Number <> 0 Then
'	  ' try legacy provider as a fallback
'	  Err.Clear
'	  If useIntegratedSecurity Then
'		connStr = "Provider=SQLOLEDB;Data Source=" & server & ";Initial Catalog=" & database & ";Integrated Security=SSPI;"
'	  Else
'		connStr = "Provider=SQLOLEDB;Data Source=" & server & ";Initial Catalog=" & database & ";User ID=" & uid & ";Password=" & pwd & ";"
'	  End If
'	  cn.Open connStr
'	  If Err.Number <> 0 Then
'		WScript.Echo "Failed to connect: " & Err.Description
'		WScript.Quit 1
'	  End If
'	End If

	' Read file into byte array
	Dim fileBytes: fileBytes = ReadAllBytes(filePath)
	If IsEmpty(fileBytes) Then
		Session.Output "Could not read file: " & filePath
		exit function
	End If

	' Prepare parameterized INSERT (ORDER OF ? PLACEHOLDERS MATTERS)
	Dim sql
	sql = "INSERT INTO t_document (DocID, DocName, ElementID, ElementType, BinContent, DocType) " & _
		  "VALUES (NEWID(), ?, ?, ?, ?, ?);"

	Dim cmd: Set cmd = CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = cn
	cmd.CommandType = adCmdText
	cmd.CommandText = sql

	' Append parameters in the same order as the ? placeholders
	cmd.Parameters.Append cmd.CreateParameter("@DocName",    adVarWChar,    adParamInput, 100, docName)
	cmd.Parameters.Append cmd.CreateParameter("@ElementID",  adVarWChar,    adParamInput,  40, elementId)
	cmd.Parameters.Append cmd.CreateParameter("@ElementType",adVarWChar,    adParamInput,  50, elementType)

	' Binary parameter (use LongVarBinary for large files)
	Dim binSize: binSize = ByteCount(fileBytes)
	Dim pBin: Set pBin = cmd.CreateParameter("@BinContent", adLongVarBinary, adParamInput, binSize)
	pBin.Value = fileBytes
	cmd.Parameters.Append pBin

	cmd.Parameters.Append cmd.CreateParameter("@DocType",    adVarWChar,    adParamInput, 100, docType)

	' Execute
	cmd.Execute

	cn.Close
	Set cn = Nothing
end function


'====================== Helpers ======================

Function ReadAllBytes(path)
  On Error Resume Next
  Dim stm: Set stm = CreateObject("ADODB.Stream")
  stm.Type = adTypeBinary
  stm.Open
  stm.LoadFromFile path
  If Err.Number <> 0 Then
    ReadAllBytes = Empty
    Exit Function
  End If
  ReadAllBytes = stm.Read
  stm.Close
  Set stm = Nothing
End Function

Function ByteCount(bytes)
  If IsEmpty(bytes) Then
    ByteCount = 0
  Else
    ByteCount = (UBound(bytes) - LBound(bytes) + 1)
  End If
End Function

main