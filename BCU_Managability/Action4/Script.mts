'*********************************************************************************************************************************************
'Name                 :  Jagadish           
'Description          :  Step 002 - Generate repset file with /Getconfig  /Advanced switch
'Pre-Requisite        :
'Created By           :      
'Creation Date        :                     
'Application Name     :                    
'Changed By         Date @ Time:            Description:                        
'**********************************************************************************************************************************************
systemutil.Run "C:\CMIT\Setup.exe"
wait 2
Window("InstallShield Wizard").WinButton("Next >").Click
Window("InstallShield Wizard").WinButton("Next >").Click
Window("InstallShield Wizard").WinButton("Install").Click
Window("InstallShield Wizard").WinButton("Finish").Click

Set fso=createobject("Scripting.FileSystemObject")
filepath="C:\Program Files (x86)\HP\BIOS Configuration Utility\BiosConfigUtility.exe"
If fso.FileExists(filepath) then
 Reporter.ReportEvent micPass, "BiosConfigUtility.exe setup Exists", ""
Else
 Reporter.ReportEvent micFail, "BiosConfigUtility.exe File doesnot Exist", ""
End If
Set fso=nothing

Set Fso = CreateObject ("Scripting.FileSystemObject")
 FilePath = "C:\TestFolder4"
 If Fso.FolderExists (FilePath) Then
 Fso.DeleteFolder (FilePath)
 Fso.CreateFolder"C:\TestFolder4"
 else
 Fso.CreateFolder"C:\TestFolder4"
 	 End If
  Set Fso = Nothing
  
'Set fso = CreateObject("scripting.FileSystemObject")
' Set T_Folder = fso.CreateFolder("C:\TestFolder4")
' Set fso=nothing

Set fso=createobject("Scripting.FileSystemObject")
fso.CopyFolder "C:\Program Files (x86)\HP\BIOS Configuration Utility","C:\TestFolder4",True
SystemUtil.Run "cmd.exe","","","runas",10
Set WshShell = CreateObject("wscript.Shell")
'WshShell.Run "cmd"
wait 2
WshShell.SendKeys "cd/"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
WshShell.SendKeys "cd TestFolder4"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
WshShell.SendKeys " BiosConfigUtility.exe /getconfig:advanced.txt"
wait 2
WshShell.SendKeys "{Enter}"
Wait 2
'Window("Administrator: C:\windows\SysW").CaptureBitmap "C:\CMIT\advanced.png"
Window("Commandpromt").CaptureBitmap "C:\CMIT\advanced.png"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
WshShell.SendKeys "BiosConfigUtility64.exe /getconfig:advanced64.txt"
Wait 2
WshShell.SendKeys "{Enter}"
wait 2
Window("Commandpromt").CaptureBitmap "C:\CMIT\advanced64.png"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
WshShell.SendKeys "Exit"
Wait 2
WshShell.SendKeys "{Enter}"

wait 2
Const FOR_READING = 1
Const FOR_WRITING = 2
strFileName = "C:\TestFolder4\advanced.txt"
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




'-------------Compare two files function-----------------------------------------------------
File1= "C:\TestFolder3\test32.txt"
File2= "C:\TestFolder4\advanced.txt"

If CompareFiles(File1, File2) = False Then
var="Files are identical"
 Set sh = CreateObject("wscript.Shell")
 sh.Popup var,"2"
Else
 var="Files are different"
  Set sh = CreateObject("wscript.Shell")
 sh.Popup var,"2"
End If

Public Function CompareFiles (FilePath1, FilePath2)
 Dim FS, File1, File2
 Set FS = CreateObject("scripting.FileSystemObject")

If FS.GetFile(FilePath1).Size <> FS.GetFile(FilePath2).Size Then
 CompareFiles = True
 Exit Function
 End If
 Set File1 = FS.GetFile(FilePath1).OpenAsTextStream(1, 0)
 Set File2 = FS.GetFile(FilePath2).OpenAsTextStream(1, 0)

CompareFiles = False
 Do While File1.AtEndOfStream = False
 Str1 = File1.Read(1000)
 Str2 = File2.Read(1000)

CompareFiles = StrComp(Str1, Str2, 0)

If CompareFiles <> 0 Then
 CompareFiles = True
 Exit Do
 End If
 Loop

File1.Close()
 File2.Close()
 End Function 
 
 '-----------Uninstall "HP BIOS Configuration Utility" through control panel----------------------------------------------


Call SW_Uninstallation ("HP BIOS Configuration Utility")
