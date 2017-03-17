'*********************************************************************************************************************************************
'Name                 :  Jagadish           
'Description          :  Step 001 - Set BIOS setting with wrong input files name
'Pre-Requisite        :
'Created By           :      
'Creation Date        :                     
'Application Name     :                    
'Changed By         Date @ Time:            Description:                        
'*******************************************************************************************************************************************
systemutil.Run "C:\CMIT\Setup.exe"
window("InstallShield Wizard").winbutton("Next >").Click
window("InstallShield Wizard").winbutton("Next >").Click
window("InstallShield Wizard").WinButton("Install").Click
window("InstallShield Wizard").WinButton("Finish").Click
'---------TestCase Step 001 - Set BIOS setting with wrong input files name-----------------------------------------
'-------Create folder if exist delete and create--------------------------------------------------------------------
Set Fso = CreateObject ("Scripting.FileSystemObject")
 FilePath = "C:\Test"
 If Fso.FolderExists (FilePath) Then
 Fso.DeleteFolder (FilePath)
  Fso.CreateFolder"C:\Test"
 else
 Fso.CreateFolder"C:\Test"
 	 End If
  Set Fso = Nothing
'---------Step1.Copy BiosConfigUtility.exe or BiosConfigUtility64.exe to filestore (e.g. c:\test)-------------------------    
Set fso=CreateObject("Scripting.FileSystemObject")
fso.CopyFolder "C:\Program Files (x86)\HP\BIOS Configuration Utility","C:\Test",True  

'--------Step2.From Start/Run type cmd------------------------------------------------------------------------------------
SystemUtil.Run "cmd.exe","","","runas",10
'WshShell.Run "cmd"
'WshShell.Run "cmd"
Set WshShell = CreateObject("WScript.Shell")
wait 2
WshShell.SendKeys "cd/"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
WshShell.SendKeys "cd test"
wait 2
WshShell.SendKeys "{Enter}"

'-------------Step3.MSDos command windows open, change to c:\test directory then Run----------------------------------------
wait 2
WshShell.SendKeys "BiosConfigUtility.exe/WarningAsErr /getconfig:test1.txt"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
Window("CommandPromt").CaptureBitmap "C:\CMIT\Wrongfiletest1.png"
wait 2
WshShell.SendKeys "{Enter}"

'------------Step4.After BCU creates test1.txt,  Run -----------------------------------------------------------------------

wait 2
WshShell.SendKeys "BiosConfigUtility.exe /WarningAsErr /setconfig:notest1.txt"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
Window("CommandPromt").CaptureBitmap "C:\CMIT\Wrongnotest1.png"

wait 2
WshShell.SendKeys "Exit"
wait 2
WshShell.SendKeys "{Enter}"
'-----------Uninstall "HP BIOS Configuration Utility" through control panel----------------------------------------------


Call SW_Uninstallation ("HP BIOS Configuration Utility")
