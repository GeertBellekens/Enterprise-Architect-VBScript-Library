'[path=\Framework\ho]
'[group=ho]

'Function:  Useful String functions
'File:      String.vbs
'Author:    Helmut Ortmann
'Date: 2015-12-30
!INC Local Scripts.EAConstants-VBScript

' Left pads a string to the specified length, truncate it if to long
Function Lpad (pString, pChar, pLength)  
  strString = pString
  
  ' Create padding of required length
  strPadding = String(pLength,pChar)  
  
  strString = strPadding & strString
  strString = Right(strString,pLength)
  
   Lpad = strString  ' Return string  

End Function 

' Right pads a string to the specified length, truncate it if to long
Function Rpad (pString, pChar, pLength)  
  strString = pString
  
  ' Create padding of required length
  strPadding = String(pLength,pChar)  
  
  strString = strString & strPadding
  strString = Left(strString,pLength)  
  
  Rpad = strString  ' Return string  

End Function  

Sub testString
Session.Output "'" + Lpad("ab","_", 10) + "'"
Session.Output "'" + Rpad("ab","_", 10) + "'"
Session.Output "'" + Lpad("1234567890123456789","_", 10) + "'"
End sub


' for test purposes remove tick mark
'testString