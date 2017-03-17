'*********************************************************************************************************************************************
'Name                 :  Jagadish           
'Description          :  Step 001 - Generate repset file with /Getconfig switch
'Pre-Requisite        :
'Created By           :      
'Creation Date        :                     
'Application Name     :                    
'Changed By         Date @ Time:            Description:                        
'**********************************************************************************************************************************************
systemutil.Run "C:\CMIT\Setup.exe"
Window("InstallShield Wizard").Activate
Window("InstallShield Wizard").WinButton("Next >").Click
Window("InstallShield Wizard").WinButton("Next >").Click
Window("InstallShield Wizard").WinButton("Install").Click
Window("InstallShield Wizard").WinButton("Finish").Click

Set fso=createobject("Scripting.FileSystemObject")
filepath="C:\Program Files (x86)\HP\BIOS Configuration Utility\BiosConfigUtility.exe"
If fso.FileExists(filepath) then
Reporter.ReportEvent micPass,"BiosConfigUtility.exe setup Exists", ""
 var = "BiosConfigUtility64.exe File Exists"
 Call generatelogs (var)
Else
 Reporter.ReportEvent micFail, "BiosConfigUtility.exe  File doesnot Exist", ""
  var = "BiosConfigUtility64.exe File doesn't  Exists"
 Call generatelogs (var)
End If
Set fso=nothing

Set Fso = CreateObject ("Scripting.FileSystemObject")
 FilePath = "C:\TestFolder3"
 If Fso.FolderExists (FilePath) Then
 Fso.DeleteFolder (FilePath)
  Fso.CreateFolder"C:\TestFolder3"
 else
 Fso.CreateFolder"C:\TestFolder3"
 	 End If
  Set Fso = Nothing

'Set fso = CreateObject("scripting.FileSystemObject")
' Set T_Folder = fso.CreateFolder("C:\TestFolder3")
' Set fso=nothing

Set fso=createobject("Scripting.FileSystemObject")
fso.CopyFolder "C:\Program Files (x86)\HP\BIOS Configuration Utility","C:\TestFolder3",True
SystemUtil.Run "cmd.exe","","","runas",10
Set WshShell = CreateObject("wscript.Shell")
'WshShell.Run "cmd"
wait 2
WshShell.SendKeys "cd/"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
WshShell.SendKeys "cd TestFolder3"
wait 2
WshShell.SendKeys "{Enter}"

Dim MyTime
MyTime = Time
time1 = MyTime
'MsgBox MyTime

wait 2
WshShell.SendKeys "BiosConfigUtility.exe /getconfig:test32.txt /log"
wait 2
WshShell.SendKeys "{Enter}"

Dim MyTime2
MyTime2 = Time
time2 = MyTime2
time3 = time1 - time2
 var = trim(time3) & "execution time in secounds"
 Call generatelogs (var)

'wait 2
'WshShell.SendKeys "BiosConfigUtility.exe /getconfig:test64.txt /log64"
'Wait 2
'WshShell.SendKeys "{Enter}"
Wait 2
WshShell.SendKeys "Exit"
Wait 2
WshShell.SendKeys "{Enter}"


wait 2
Const FOR_READING = 1
Const FOR_WRITING = 2
strFileName = "C:\TestFolder3\test32.txt"
iNumberOfLinesToDelete = 5

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(strFileName, FOR_READING)
strContents = objTS.ReadAll
objTS.Close

arrLines = Split(strContents, vbNewLine)
Set objTS = objFS.OpenTextFile(strFileName, FOR_WRITING)

For i=0 To UBound(arrLines)
If i >(iNumberOfLinesToDelete - 1) Then
 objTS.WriteLine arrLines(i)
End If
Next

Set fso=createobject("Scripting.FileSystemObject")
filepath="C:\TestFolder3\test32.txt"
If fso.FileExists(filepath) then
Reporter.ReportEvent micPass,"Generate test32.txt file Exists", ""
Else
 Reporter.ReportEvent micFail, "Generate test32.txt file doesn't Exist", ""
End If
Set fso=nothing


Set fso=createobject("Scripting.FileSystemObject")
filepath="C:\TestFolder3\Logs\BiosConfigUtility-0.log"
If fso.FileExists(filepath) then
Reporter.ReportEvent micPass,"Logs file,Folder  Exists", ""
 var = "Logs file,Folder  Exists"
 Call generatelogs (var)
Else
 Reporter.ReportEvent micFail, "Logs file,Folder doesn't Exist", ""
  var = "Logs file,Folder doesn't Exist"
 Call generatelogs (var)
End If
Set fso=nothing

Set fso=createobject("Scripting.FileSystemObject")
filepath="C:\TestFolder3\Logs\BiosConfigUtility-0.log"
If fso.FileExists(filepath) then
Reporter.ReportEvent micPass,"BiosConfigUtility-0.log   Exists", ""
  var = "BiosConfigUtility-0.log   Exists"
 Call generatelogs (var)
Else
 Reporter.ReportEvent micFail, "BiosConfigUtility-0.log  doesn't Exist", ""
   var = "BiosConfigUtility-0.log doesn't  Exists"
 Call generatelogs (var)
End If
Set fso=nothing


Set objFSO = CreateObject("Scripting.FileSystemObject")
wait 3
Set objFile = objFSO.OpenTextFile("C:\TestFolder3\Logs\BiosConfigUtility-0.log",1) ' Open in read mode
wait 3

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
      if Instr(strLine, "SUCCESS") > 1 then 
       f=1
   End IF	    
loop 
If f = 1 Then
	     	Reporter.ReportEvent micPass, "Successfully installed", ""
	     	else
	     	Reporter.ReportEvent micFail, "Successfully not installed", ""
	       End If
objFile.Close

'-----------Uninstall "HP BIOS Configuration Utility" through control panel----------------------------------------------


Call SW_Uninstallation ("HP BIOS Configuration Utility")




