'*********************************************************************************************************************************************
'Name                 :  Jagadish           
'Description          :  Set BIOS Setting with unchanged setting value
'Pre-Requisite        :  Need to copy test3.txt file in the path "C:\Test"
'Created By           :      
'Creation Date        :                     
'Application Name     :                    
'Changed By         Date @ Time:            Description:                        
'********************************************************************************************************************************************** 
systemutil.Run "C:\CMIT\Setup.exe"
wait 10
window("InstallShield Wizard").winbutton("Next >").Click
window("InstallShield Wizard").winbutton("Next >").Click
'window("InstallShield Wizard").winbutton("Next >").Click
window("InstallShield Wizard").WinButton("Install").Click
window("InstallShield Wizard").WinButton("Finish").Click

'systemutil.Run "C:\CMIT\Permission.bat"

SystemUtil.Run "cmd.exe","","","runas",10
Set WshShell = CreateObject("wscript.Shell")
wait 2
WshShell.Sendkeys "icacls"
'WshShell.Sendkeys """"
WshShell.Sendkeys " """
WshShell.Sendkeys "C:\Program Files" 
WshShell.SendKeys" +9"
WshShell.SendKeys"x86"
WshShell.SendKeys"+0"
WshShell.SendKeys"\HP\BIOS Configuration Utility"
WshShell.Sendkeys """"
WshShell.Sendkeys " /grant Users:"
WshShell.SendKeys"+9"
WshShell.Sendkeys "OI"
WshShell.SendKeys"+0"
WshShell.SendKeys"+9"
WshShell.SendKeys "CI"
WshShell.SendKeys"+0"
WshShell.SendKeys "F /T"
'WshShell.Sendkeys """"
WshShell.SendKeys "{Enter}"
wait 2
WshShell.SendKeys "Exit"
WshShell.SendKeys "{Enter}"

Set Fso = CreateObject ("Scripting.FileSystemObject")
 FilePath = "C:\Test"
 If Fso.FolderExists (FilePath) Then
 Fso.DeleteFolder (FilePath)
 'msgbox 1
  Fso.CreateFolder"C:\Test"
 else
 'msgbox 2
 Fso.CreateFolder"C:\Test"
 	 End If
  Set Fso = Nothing
  
 
'-----------------------------1.Copy BiosConfigUItility.exe  or BiosConfigUtility64.exe  to filestore (e.g. c:\test)----------------  
Set fso=CreateObject("Scripting.FileSystemObject")
fso.CopyFolder "C:\Program Files (x86)\HP\BIOS Configuration Utility","C:\Test",True 

'----------------------------2.From Start/Run type cmd-----------------------------------
SystemUtil.Run "cmd.exe","","","runas",10
Set WshShell = CreateObject("wscript.Shell")
'WshShell.Run "cmd"
wait 2
WshShell.SendKeys "cd/"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
WshShell.SendKeys "cd C:\Test"
wait 2
WshShell.SendKeys "{Enter}"
'-------------3. MSDos command windows open, change to c:\test directory then Run "BiosConfigUtility.exe /getconfig:test1.txt"--------
wait 2
WshShell.SendKeys "BiosConfigUtility.exe /getconfig:test1.txt"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
Window("Commandpromt").CaptureBitmap "C:\CMIT\getconfigunchangedsetting.png"

wait 2
WshShell.SendKeys "Exit"
wait 2
WshShell.SendKeys "{Enter}"
'------------4. After BCU creates test1.txt, open and pick a string setting (Asset tag) but do not change its value.------------------
'-----------Asset tag doesn't exits in the generated file-----------------------------------------------------------------------------
'-----------5. Clear all other settings.  Save test3.txt------------------------------------------------------

Dim fso 
Set fso = CreateObject("Scripting.FileSystemObject")
Dim oFile 
StrPath= "C:\Test\test3.txt"
Set oFile = FSO.CreateTextFile(strPath)
'oFile.WriteLine "test" 
oFile.Close
Set fso = Nothing
Set oFile = Nothing 

Set objFSO = CreateObject("Scripting.FileSystemObject")
wait 3
Set objFile = objFSO.OpenTextFile("C:\Test\test1.txt",1) ' Open in read mode
wait 3

Set objFile1 = objFSO.OpenTextFile("C:\Test\test3.txt",2)' Open in write mode

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
	     spaces=InStr(strLine, "	")       
	    
	   If (spaces = 1) Then   
		  	strLine=""
	   End If

	objFile1.WriteLine(strLine)

loop 

objFile.Close
objFile1.Close

'-----------6. Run:BiosConfigUtility.exe /WarningAsErr /setconfig:test1. txt /verbose"--------------------
systemutil.Run "C:\CMIT\Permission1.bat"
Set fso=CreateObject("Scripting.FileSystemObject")
fso.CopyFolder "C:\CMIT","C:\Test",True  
Set Fso = Nothing 
SystemUtil.Run "cmd.exe","","","runas",10
Set WshShell = CreateObject("wscript.Shell")
wait 2
WshShell.SendKeys "cd/"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
WshShell.SendKeys "cd C:\Test"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
WshShell.SendKeys "BiosConfigUtility.exe /WarningAsErr /setconfig:test3.txt"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
Window("Commandpromt").CaptureBitmap "C:\CMIT\Setconfigunchangedsetting.png"
wait 2
WshShell.SendKeys "Exit"
wait 2
WshShell.SendKeys "{Enter}"

'-----------Uninstall "HP BIOS Configuration Utility" through control panel----------------------------------------------

Call SW_Uninstallation ("HP BIOS Configuration Utility")
