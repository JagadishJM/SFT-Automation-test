'*********************************************************************************************************************************************
'Name                 :  Jagadish           
'Description          :  Step 001 - Set Password to 8 character including letters and numbers from the top of keyboard
'Pre-Requisite        :
'Created By           :      
'Creation Date        :                     
'Application Name     :                    
'Changed By         Date @ Time:            Description:                        
'**********************************************************************************************************************************************
'------------------------Step 001 - Set Password to 8 character including letters and numbers from the top of keyboard---------------------------------------
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
     var = "BiosConfigUtility.exe setup Exists"
   Call generatelogs (var)
 Else
 Reporter.ReportEvent micFail, "BiosConfigUtility.exe setup File doesnot Exist", ""
  var = "BiosConfigUtility.exe setup File doesnot Exist"
   Call generatelogs (var)
 End If
Set fso=nothing

Set Fso = CreateObject ("Scripting.FileSystemObject")
 FilePath = "C:\TestFolder5"
 If Fso.FolderExists (FilePath) Then
 Fso.DeleteFolder (FilePath)
 Fso.CreateFolder"C:\TestFolder5"
 else
 Fso.CreateFolder"C:\TestFolder5"
 	 End If
  Set Fso = Nothing

Set fso=createobject("Scripting.FileSystemObject")
fso.CopyFolder "C:\Program Files (x86)\HP\BIOS Configuration Utility","C:\TestFolder5",True

SystemUtil.Run "cmd.exe","","","runas",10
Set WshShell = CreateObject("wscript.Shell")
'WshShell.Run "cmd"
wait 2
WshShell.SendKeys "cd/"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
WshShell.SendKeys "cd C:\TestFolder5"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
WshShell.SendKeys "HPQPswd.exe"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
Dialog("HPQPswd 2013").WinEdit("Password to be encrypted").Click
Dialog("HPQPswd 2013").WinEdit("Password to be encrypted").Type "HP123456.bin"
WshShell.SendKeys "{TAB}"
Dialog("HPQPswd 2013").WinEdit("Re-enter Password").Click
Dialog("HPQPswd 2013").WinEdit("Re-enter Password").Type "HP123456.bin"
WshShell.SendKeys "{TAB}"
wait 2
Dialog("HPQPswd 2013").WinEdit("File to save encrypted").Type "C:\TestFolder5\HP123456.bin"
wait 2
Dialog("HPQPswd 2013").WinButton("OK").Click

wait 2
WshShell.SendKeys "BiosConfigUtility.exe /WarningAsErr /nspwdfile:"
WshShell.SendKeys """"
WshShell.SendKeys "HP123456.bin"
WshShell.SendKeys """"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
WshShell.SendKeys "Exit"
wait 2
WshShell.SendKeys"{Enter}"

'-----------Uninstall "HP BIOS Configuration Utility" through control panel----------------------------------------------


Call SW_Uninstallation ("HP BIOS Configuration Utility")
