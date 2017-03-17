'*********************************************************************************************************************************************
'Name                 :  Jagadish           
'Description          :  Step 001 - Short Command line /Get = /Getconfig
'Pre-Requisite        :
'Created By           :      
'Creation Date        :                     
'Application Name     :                    
'Changed By         Date @ Time:            Description:                        
'**********************************************************************************************************************************************
'--------------Test Steps Number [009]------------------------
'-----------------------------Step 001 - Short Command line /Get = /Getconfig---------------------------

systemutil.Run "C:\CMIT\Setup.exe"
window("InstallShield Wizard").winbutton("Next >").Click
window("InstallShield Wizard").winbutton("Next >").Click
'window("InstallShield Wizard").winbutton("Next >").Click
window("InstallShield Wizard").WinButton("Install").Click
window("InstallShield Wizard").WinButton("Finish").Click

'systemutil.Run "C:\CMIT\Permission.bat"

'SystemUtil.Run "cmd.exe","","","runas",10
'Set WshShell = CreateObject("wscript.Shell")
'wait 2
'WshShell.Sendkeys "icacls"
''WshShell.Sendkeys """"
'WshShell.Sendkeys " """
'WshShell.Sendkeys "C:\Test" 
'WshShell.SendKeys" +9"
'WshShell.SendKeys"x86"
'WshShell.SendKeys"+0"
'WshShell.SendKeys"\HP\BIOS Configuration Utility"
'WshShell.Sendkeys """"
'WshShell.Sendkeys " /grant Users:"
'WshShell.SendKeys"+9"
'WshShell.Sendkeys "OI"
'WshShell.SendKeys"+0"
'WshShell.SendKeys"+9"
'WshShell.SendKeys "CI"
'WshShell.SendKeys"+0"
'WshShell.SendKeys "F /T"
'WshShell.Sendkeys """"
'WshShell.SendKeys "{Enter}"
'wait 2
'WshShell.SendKeys "Exit"
'WshShell.SendKeys "{Enter}"

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
'-----------------------------2.Copy BiosConfigUItility.exe  or BiosConfigUtility64.exe  to filestore (e.g. c:\test)---------------- 
 
Set fso=CreateObject("Scripting.FileSystemObject")
fso.CopyFolder "C:\Program Files (x86)\HP\BIOS Configuration Utility","C:\Test",True  
Set Fso = Nothing   

'-------------------------3. MsDos command windows open,  change to  c:\test directory then Run -------------------------
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
WshShell.SendKeys "BiosConfigUtility.exe /Get:Test123.txt"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
Window("Commandpromt").CaptureBitmap "C:\CMIT\Getconfig.png"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
WshShell.SendKeys "Exit"
wait 2
WshShell.SendKeys "{Enter}"

'-----------Uninstall "HP BIOS Configuration Utility" through control panel----------------------------------------------

Call SW_Uninstallation ("HP BIOS Configuration Utility")
