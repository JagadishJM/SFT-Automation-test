﻿'*********************************************************************************************************************************************
'Name                 :  Jagadish           
'Description          :  Step 006 - Set BIOS Setting with wrong setting name
'Pre-Requisite        :  
'Created By           :      
'Creation Date        :                     
'Application Name     :                    
'Changed By         Date @ Time:            Description:                        
'**********************************************************************************************************************************************

systemutil.Run "C:\CMIT\Setup.exe"
window("InstallShield Wizard").winbutton("Next >").Click
window("InstallShield Wizard").winbutton("Next >").Click
'window("InstallShield Wizard").winbutton("Next >").Click
window("InstallShield Wizard").WinButton("Install").Click
window("InstallShield Wizard").WinButton("Finish").Click

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
Set Fso = Nothing 
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

'-------------3. MSDos command windows open, change to c:\test directory then Run BiosConfigUtility.exe /getconfig:test1.txt--------
wait 2
WshShell.SendKeys "BiosConfigUtility.exe /getconfig:test1.txt"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
Window("Commandpromt").CaptureBitmap "C:\CMIT\test1Biossettingwrongname.png"
wait 2
WshShell.SendKeys "Exit"
wait 2
WshShell.SendKeys "{Enter}"

'----------4. After BCU creates test1.txt, open it and change a setting name to a non-existant name (e.g. change 'Asset Tag' to 'Asset Tug'  or change 'Asset Tracking number' to 'Asset Track Num')--
'Dim fso 
Set fso = CreateObject("Scripting.FileSystemObject")
'Dim oFile 
StrPath= "C:\Test\test3.txt"
Set oFile = FSO.CreateTextFile(strPath)
'oFile.WriteLine "test" 
oFile.Close
Set fso = Nothing
Set oFile = Nothing 

Dim fileContent,fso
Dim strText
'Dim fso,fileContent
Set fso= CreateObject("Scripting.FileSystemObject")
wait 2
Set fileContent =fso.OpenTextFile("C:\Test\test1.txt",1)
strText=fileContent.ReadAll
fileContent.Close

strText=Replace (strText,"Asset Tracking Number","Asset Track Num")

Set fileContent =fso.OpenTextFile("C:\Test\test1.txt",2)
fileContent.writeLine(strText)
fileContent.Close

'------------------------5. Clear all other settings.  Save test1.txt-----------------------------------------------------------
Set objFSO = CreateObject("Scripting.FileSystemObject")
Wait 3
Set objFile = objFSO.OpenTextFile("C:\Test\test1.txt",1) ' Open in read mode
wait 2
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

'------------------------6. Run:BiosConfigUtility.exe /WarningAsErr /setconfig:test1. txt /verbose"-----------------------------
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
wait 2
WshShell.SendKeys "BiosConfigUtility.exe /WarningAsErr /setconfig:test3.txt /verbose"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
Window("Commandpromt").CaptureBitmap "C:\CMIT\test3verbosewrongsettingname.png"
wait 2
WshShell.SendKeys "Exit"
wait 2
WshShell.SendKeys "{Enter}"

'-----------Uninstall "HP BIOS Configuration Utility" through control panel----------------------------------------------


Call SW_Uninstallation ("HP BIOS Configuration Utility")
