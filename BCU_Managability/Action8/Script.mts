'*********************************************************************************************************************************************
'Name                 :  Jagadish           
'Description          :  Step 002 - Set BIOS setting with no 'BIOSConfig' or 'English' as first line
'Pre-Requisite        :
'Created By           :      
'Creation Date        :                     
'Application Name     :                    
'Changed By         Date @ Time:            Description:                        
'**********************************************************************************************************************************************
'----------------TestCase Step 002 - Set BIOS setting with no 'BIOSConfig' or 'English' as first line -----------------------------------------------
'-------Create folder if exist delete and create-----------------------------------------------------------------------------------------------------
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
  Fso.CreateFolder"C:\Test"
 else
 Fso.CreateFolder"C:\Test"
 	 End If
  Set Fso = Nothing
 
Set fso=CreateObject("Scripting.FileSystemObject")
fso.CopyFolder "C:\Program Files (x86)\HP\BIOS Configuration Utility","C:\Test",True  
Set Fso = Nothing 
  
'--------Step2.From Start/Run type cmd------------------------------------------------------------------------------------
SystemUtil.Run "cmd.exe","","","runas",10
Set WshShell = CreateObject("wscript.Shell")
'WshShell.Run "cmd"
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
wait 2
Window("Commandpromt").CaptureBitmap "C:\CMIT\BIOSsettingwithnoBIOSConfigorEnglishasfirstline1.png"

'------------Step 4.BCU creates test1.txt, open it and remove the first line containing either 'BIOSConfig' or 'English' & 5. Save test1.txt -------------
wait 2
Const FOR_READING = 1
Const FOR_WRITING = 2
strFileName = "C:\Test\test1.txt"
iNumberOfLinesToDelete = 1

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

'-----------Step 6. Run  the following MSDos from opened MSDos windows BiosConfigUtility.exe /WarningAsErr/setconfig:test1.txt /verbose"--------------------------

wait 2
WshShell.SendKeys "BiosConfigUtility.exe /WarningAsErr /setconfig:notest1.txt /verbose"
wait 2
WshShell.SendKeys "{Enter}"
wait 2
Window("Commandpromt").CaptureBitmap "C:\CMIT\BIOSsettingwithnoBIOSConfigorEnglishasfirstline2.png"
wait 2
WshShell.SendKeys "Exit"
wait 2
WshShell.SendKeys "{Enter}"

 '-----------Uninstall "HP BIOS Configuration Utility" through control panel----------------------------------------------


Call SW_Uninstallation ("HP BIOS Configuration Utility")
