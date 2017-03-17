'---------------Step 006 - Two nspwdfile parameters in the same command line-------------------------------------------------------------
'-------------1 .Ensure there is no password set in F10--------------------------------------

systemutil.Run "C:\CMIT\Setup.exe"
window("InstallShield Wizard").winbutton("Next >").Click
'window("InstallShield Wizard").winbutton("Next >").Click
window("InstallShield Wizard").winbutton("Next >").Click
window("InstallShield Wizard").WinButton("Install").Click
window("InstallShield Wizard").WinButton("Finish").Click


Set Fso = CreateObject ("Scripting.FileSystemObject")
 FilePath = "C:\Test"
 If Fso.FolderExists (FilePath) Then
 Fso.DeleteFolder (FilePath)

  Fso.CreateFolder"C:\Test"
 else

 Fso.CreateFolder"C:\Test"
 	 End If
  Set Fso = Nothing
'-----------------------------2.Copy BiosConfigUItility.exe  or BiosConfigUtility64.exe  to filestore (e.g. c:\test)---------------- 
 
Set fso=createobject("Scripting.FileSystemObject")
fso.CopyFolder "C:\Program Files (x86)\HP\BIOS Configuration Utility","C:\Test",True  
Set Fso = Nothing 

'----------------------------3. From Start/Run type cmd------------------------------------------------------------------------------
'------------------create 1st encrypted password file that has 8 char password ""P@ssword""---------------------------------------
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
WshShell.SendKeys "HPQPswd.exe"
wait 2
WshShell.SendKeys "{Enter}"

wait 2
Dialog("HPQPswd 2013").WinEdit("Password to be encrypted").Click
Dialog("HPQPswd 2013").WinEdit("Password to be encrypted").Type "P@ssword"
wait 2
Dialog("HPQPswd 2013").WinEdit("Password to be encrypted").Click
Dialog("HPQPswd 2013").WinEdit("Re-enter Password").Type "P@ssword"
wait 2
Dialog("HPQPswd 2013").WinEdit("File to save encrypted").Type "C:\Test\P@ssword.bin"
'Dialog("HPQPswd 2013").WinButton("Browse").Click
'wait 2
'Dialog("Open").WinObject("Items View").WinList("Items View").Select "P@ssword.bin"
'wait 2
'Dialog("Open").WinButton("Open").Click
wait 2
Dialog("HPQPswd 2013").WinButton("OK").Click
'Dialog("HPQPswd").WinButton("Yes").Click
'Dialog("HPQPswd").WinButton("OK").Click
wait 2
WshShell.SendKeys "exit"
wait 2
WshShell.SendKeys "{Enter}"
'---------create the 2nd encrypted password file that has 32 char ""TestBCU@HP123456789abcd1234<>789"" and name it to (e.g. new32combo.bin)----
'-------------5. Save password bin files to BCU filestore (e.g. c:\test)---------------------------------------
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
WshShell.SendKeys "HPQPswd.exe"
wait 2
WshShell.SendKeys "{Enter}"

wait 2
Dialog("HPQPswd 2013").WinEdit("Password to be encrypted").Click
Dialog("HPQPswd 2013").WinEdit("Password to be encrypted").Type "TestBCU@HP123456789abcd1234<>789"
wait 2
Dialog("HPQPswd 2013").WinEdit("Password to be encrypted").Click
Dialog("HPQPswd 2013").WinEdit("Re-enter Password").Type "TestBCU@HP123456789abcd1234<>789"
wait 2
Dialog("HPQPswd 2013").WinEdit("File to save encrypted").Type "C:\Test\new32combo.bin"
'Dialog("HPQPswd 2013").WinButton("Browse").Click
'wait 2
'Dialog("Open").WinObject("Items View").WinList("Items View").Select "P@ssword.bin"
'wait 2
'Dialog("Open").WinButton("Open").Click
wait 2
Dialog("HPQPswd 2013").WinButton("OK").Click
'Dialog("HPQPswd").WinButton("Yes").Click
'Dialog("HPQPswd").WinButton("OK").Click

wait 2
WshShell.SendKeys "exit"
wait 2
WshShell.SendKeys "{Enter}"

'-------------------------6. MSDos command windows open, and to c:\test directory then  Run---------------------
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
WshShell.SendKeys "BiosConfigUtility.exe /WarningAsErr /nspwdfile:""P@ssword.bin"""
wait 2
WshShell.SendKeys " /nspwdfile:""new32combo.bin"" "
wait 2
WshShell.SendKeys "{Enter}"
Wait 2
WshShell.SendKeys "exit"
wait 2
WshShell.SendKeys "{Enter}"
wait 2

'-----------Uninstall "HP BIOS Configuration Utility" through control panel----------------------------------------------

Call SW_Uninstallation ("HP BIOS Configuration Utility")

