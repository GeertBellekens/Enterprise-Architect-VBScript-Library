'[path=\Projects\Project DL\DL Scripts]
'[group=De Lijn Scripts]
option explicit 
 
' 
' Script Name: Automated backup of EA project databases
' Author: Davy Glerum, Geert Bellekens, Tom Geerts, Alain Van Goethem
' Purpose: Automated Project Transfer from DBMS to EAP file as weekly backup. See end of script for different databases that are backed up.
' Date: 07/11/2018

'
sub DeLijnDEV

 Dim CurrentDate
 Currentdate = (Year(Date) & (Right(String(2,"0") & Month(Date), 2)) & (Right(String(2,"0") & Day(Date), 2)))  'yyyymmdd'

 'dim repository
 dim projectInterface
 'set repository = CreateObject("EA.Repository")

 Dim FileName
 Filename = "EA_Export.eap"
  
 dim LogFilePath
 LogFilePath = "H:\backups\"&CurrentDate & " DeLijnDEV (back-up).log"

 dim TargetFilePath
 TargetFilePath = "H:\Backups\"&CurrentDate & " DeLijnDEV (back-up).eapx"

 dim eapString
 eapString = "DBType=3;Connect=Provider=OraOLEDB.Oracle.1;Password=SDu_udr_7rfFErtsw99_a4f;Persist Security Info=True;User ID=enterprisearchitectuurdev;Data Source=caoracdcmvvm-scan.vvm.addelijn.be:1551/DSHARED.delijn.be;"

 'get project interface
 set projectInterface = Repository.GetProjectInterface()

 projectInterface.ProjectTransfer eapString, TargetFilePath, LogFilePath

 'repository.CloseFile
 'repository.Exit 

' Dim newFilename 
' newFilename = "H:\Backups\"&CurrentDate & " DeLijnDEV (back-up).eap"
'
' Dim Fso
' Set Fso = WScript.CreateObject("Scripting.FileSystemObject")
' Fso.MoveFile TargetFilePath, newFileName
   
end sub



DeLijnDEV
'DeLijnTEST
'DeLijnPROD

MsgBox ("Back-up Finished.")