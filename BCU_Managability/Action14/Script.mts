'*********************************************************************************************************************************************
'Name                 :  Jagadish           
'Description          :  Step 1 Verify the user guide
'Pre-Requisite        :  
'Created By           :      
'Creation Date        :                     
'Application Name     :                    
'Changed By         Date @ Time:            Description:                        
'**********************************************************************************************************************************************
systemutil.Run "C:\CMIT\Setup.exe"
wait 2
window("InstallShield Wizard").winbutton("Next >").Click
window("InstallShield Wizard").winbutton("Next >").Click
'window("InstallShield Wizard").winbutton("Next >").Click
window("InstallShield Wizard").WinButton("Install").Click
window("InstallShield Wizard").WinButton("Finish").Click

'Validations after install BCU some files created that need to check
Set fso=createobject("Scripting.FileSystemObject")
'The file to check the existence
filepath="C:\Program Files (x86)\HP\BIOS Configuration Utility\BIOS Configuration Utility User's Guide.url"
If fso.FileExists(filepath) then
    Reporter.ReportEvent micPass, "BIOS Configuration Utility User's Guide.url File Exists", ""
 Else
   Reporter.ReportEvent micFail , "BIOS Configuration Utility User's Guide.url File Exists File doesnot Exist", ""
End If
Set fso=nothing

'Systemutil.Run "C:\Program Files (x86)\HP\BIOS Configuration Utility\BIOS Configuration Utility User's Guide.url"

'-----------Uninstall "HP BIOS Configuration Utility" through control panel----------------------------------------------

Call SW_Uninstallation ("HP BIOS Configuration Utility")
