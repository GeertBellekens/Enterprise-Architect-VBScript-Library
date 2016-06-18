'[path=\Framework\Utils]
'[group=Utils]

'
' Script Name:  Base64
' Author: Geert Bellekens after Antonin Foller (copied from a Stackoverflow answer)
' Purpose: Encode and Decode Base64 strings
' Date: 2016-06-11
'
Function Base64Encode(sText)
    Dim oXML, oNode

    Set oXML = CreateObject("MSXML2.DOMDocument")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.nodeTypedValue =Stream_StringToBinary(sText)
    Base64Encode = oNode.text
    Set oNode = Nothing
    Set oXML = Nothing
End Function

Function Base64Decode(ByVal vCode)
	Dim oXML, oNode
	'remove white spaces, If any
	vCode = Replace(vCode, vbCrLf, "")
	vCode = Replace(vCode, vbTab, "")
	vCode = Replace(vCode, " ", "")
	Set oXML = CreateObject("MSXML2.DOMDocument")
	Set oNode = oXML.CreateElement("base64")
	oNode.dataType = "bin.base64"
	oNode.text = vCode
	Base64Decode = Stream_BinaryToString(oNode.nodeTypedValue)
	Set oNode = Nothing
	Set oXML = Nothing
End Function

'Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data
Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.CharSet = "ascii-us"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText Text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function

'Stream_BinaryToString Function
'2003 Antonin Foller, http://www.motobit.com
'Binary - VT_UI1 | VT_ARRAY data To convert To a string 
Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write binary data To the object
  BinaryStream.Open
  BinaryStream.Write Binary

  'Change stream type To text/string
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText

  'Specify charset For the output text (unicode) data.
   BinaryStream.CharSet = "utf-8"

  'Open the stream And get text/string data from the object
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function

'Public Function Base64Decode(sString)
'
'    Dim bOut , bIn , bTrans(255) , lPowers6(63) , lPowers12(63) 
'    Dim lPowers18(63) , lQuad , iPad , lChar, lPos, sOut
'    Dim lTemp 
'
'    sString = Replace(sString, vbCr, vbNullString)      'Get rid of the vbCrLfs.  These could be in...
'    sString = Replace(sString, vbLf, vbNullString)      'either order.
'
'    lTemp = Len(sString) Mod 4                          'Test for valid input.
'    If lTemp Then
'        msgbox "Input string is not valid Base64."
'    End If
'
'    If InStrRev(sString, "==") Then                     'InStrRev is faster when you know it's at the end.
'        iPad = 2                                        'Note:  These translate to 0, so you can leave them...
'    ElseIf InStrRev(sString, "=") Then                  'in the string and just resize the output.
'        iPad = 1
'    End If
'
'    For lTemp = 0 To 255                                'Fill the translation table.
'        Select Case lTemp
'            Case 65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95
'                bTrans(lTemp) = lTemp - 65              'A - Z
'            Case 97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122
'                bTrans(lTemp) = lTemp - 71              'a - z
'            Case 48,49,50,51,52,53,54,55,56,57
'                bTrans(lTemp) = lTemp + 4               '1 - 0
'            Case 43
'                bTrans(lTemp) = 62                      'Chr(43) = "+"
'            Case 47
'                bTrans(lTemp) = 63                      'Chr(47) = "/"
'        End Select
'    Next 
'
'    For lTemp = 0 To 63                                 'Fill the 2^6, 2^12, and 2^18 lookup tables.
'        lPowers6(lTemp) = lTemp * cl2Exp6
'        lPowers12(lTemp) = lTemp * cl2Exp12
'        lPowers18(lTemp) = lTemp * cl2Exp18
'    Next 
'
'    bIn = StrConv(sString, vbFromUnicode)               'Load the input byte array.
'    ReDim bOut((((UBound(bIn) + 1) \ 4) * 3) - 1)       'Prepare the output buffer.
'
'    For lChar = 0 To UBound(bIn) Step 4
'        lQuad = lPowers18(bTrans(bIn(lChar))) + lPowers12(bTrans(bIn(lChar + 1))) + _
'                lPowers6(bTrans(bIn(lChar + 2))) + bTrans(bIn(lChar + 3))           'Rebuild the bits.
'        lTemp = lQuad And clHighMask                    'Mask for the first byte
'        bOut(lPos) = lTemp \ cl2Exp16                   'Shift it down
'        lTemp = lQuad And clMidMask                     'Mask for the second byte
'        bOut(lPos + 1) = lTemp \ cl2Exp8                'Shift it down
'        bOut(lPos + 2) = lQuad And clLowMask            'Mask for the third byte
'        lPos = lPos + 3
'    Next 
'
'    sOut = StrConv(bOut, vbUnicode)                     'Convert back to a string.
'    If iPad Then sOut = Left(sOut, Len(sOut) - iPad)   'Chop off any extra bytes.
'    Decode64 = sOut
'
'End Function

' Decodes a base-64 encoded string (BSTR type).
' 1999 - 2004 Antonin Foller, http://www.motobit.com
' 1.01 - solves problem with Access And 'Compare Database' (InStr)
'Function Base64Decode(ByVal base64String)
'  'rfc1521
'  '1999 Antonin Foller, Motobit Software, http://Motobit.cz
'  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
'  Dim dataLength, sOut, groupBegin
'  
'  'remove white spaces, If any
'  base64String = Replace(base64String, vbCrLf, "")
'  base64String = Replace(base64String, vbTab, "")
'  base64String = Replace(base64String, " ", "")
'  
'  'The source must consists from groups with Len of 4 chars
'  dataLength = Len(base64String)
''  If dataLength Mod 4 <> 0 Then
''    Err.Raise 1, "Base64Decode", "Bad Base64 string."
''    Exit Function
''  End If
'
'  
'  ' Now decode each group:
'  For groupBegin = 1 To dataLength Step 4
'    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
'    ' Each data group encodes up To 3 actual bytes.
'    numDataBytes = 3
'    nGroup = 0
'
'    For CharCounter = 0 To 3
'      ' Convert each character into 6 bits of data, And add it To
'      ' an integer For temporary storage.  If a character is a '=', there
'      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
'      ' the whole string.)
'
'      thisChar = Mid(base64String, groupBegin + CharCounter, 1)
'
'      If thisChar = "=" Then
'        numDataBytes = numDataBytes - 1
'        thisData = 0
'      Else
'        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
'      End If
'      If thisData = -1 Then
'        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
'        Exit Function
'      End If
'
'      nGroup = 64 * nGroup + thisData
'    Next
'    
'    'Hex splits the long To 6 groups with 4 bits
'    nGroup = Hex(nGroup)
'    
'    'Add leading zeros
'    nGroup = String(6 - Len(nGroup), "0") & nGroup
'    
'    'Convert the 3 byte hex integer (6 chars) To 3 characters
'    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
'      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
'      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
'    
'    'add numDataBytes characters To out string
'    sOut = sOut & Left(pOut, numDataBytes)
'  Next
'
'  Base64Decode = sOut
'End Function