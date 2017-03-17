
'*********************************************************************************************************************************************
'Name                 :  Jagadish           
'Description          :  Test Steps Number [0010] Step 001 - Help /?
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
wait 2
WshShell.SendKeys "cd/"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
WshShell.SendKeys "cd C:\Test"
wait 2
WshShell.SendKeys "{Enter}"

'-------------3. MSDos command windows open, change to c:\test directory then Run BiosConfigUtility.exe /?--------
wait 2
WshShell.SendKeys "BiosConfigUtility.exe /?"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
Window("Commandpromt").CaptureBitmap "C:\CMIT\helpcontent.png"

wait 2
WshShell.SendKeys "Exit"
wait 2
WshShell.SendKeys "{Enter}"

'-----------Uninstall "HP BIOS Configuration Utility" through control panel----------------------------------------------

Call SW_Uninstallation ("HP BIOS Configuration Utility")


