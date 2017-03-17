'*********************************************************************************************************************************************
'Name                 :  Jagadish           
'Description          :  Step 001 - Repair from BCU installation removal
'Pre-Requisite        :  BCU Softpaq need to install before start
'Created By           :      
'Creation Date        :                     
'Application Name     :                    
'Changed By         Date @ Time:            Description:                        
'**********************************************************************************************************************************************
'Validations after install BCU some files created that need to check
Set fso=createobject("Scripting.FileSystemObject")
'The file to check the existence
filepath="C:\Program Files (x86)\HP\BIOS Configuration Utility\BIOS Configuration Utility User's Guide.url"
If fso.FileExists(filepath) then
    Reporter.ReportEvent micPass, "BIOS Configuration Utility User's Guide.url File Exists", ""
    var = "BIOS Configuration Utility User's Guide.url File Exists"
   Call generatelogs (var) 
 Else
   Reporter.ReportEvent micFail , "BIOS Configuration Utility User's Guide.url File Exists File doesnot Exist", ""
    var = "BIOS Configuration Utility User's Guide.url File Exists File doesnot Exist"
    Call generatelogs (var)
End If
Set fso=nothing

Set fso1=createobject("Scripting.FileSystemObject")
filepath="C:\Program Files (x86)\HP\BIOS Configuration Utility\HPQPswd.exe"
fso1.DeleteFile("C:\Program Files (x86)\HP\BIOS Configuration Utility\HPQPswd.exe")

systemutil.Run "C:\CMIT\Setup.exe"
Window("InstallShield Wizard").WinButton("Next >").Click
If Window("InstallShield Wizard").Exist Then
	Window("InstallShield Wizard").WinRadioButton("Repair").Click
End If
Window("InstallShield Wizard").WinButton("Next >").Click
Window("InstallShield Wizard").WinButton("Install").Click
'Window("InstallShield Wizard").Dialog("User Account Control").WinButton("Yes").Click
Window("InstallShield Wizard").WinButton("Finish").Click

Set fso=createobject("Scripting.FileSystemObject")
'The file to check the existence
filepath="C:\Program Files (x86)\HP\BIOS Configuration Utility\BIOS Configuration Utility User's Guide.url"
If fso.FileExists(filepath) then
    Reporter.ReportEvent micPass, "BIOS Configuration Utility User's Guide.url File Exists", ""
    var = "BIOS Configuration Utility User's Guide.url File Exists"
   Call generatelogs (var) 
 Else
   Reporter.ReportEvent micFail , "BIOS Configuration Utility User's Guide.url File Exists File doesnot Exist", ""
    var = "BIOS Configuration Utility User's Guide.url File Exists File doesnot Exist"
    Call generatelogs (var)
End If
Set fso=nothing

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

Set fso=createobject("Scripting.FileSystemObject")
filepath="C:\Program Files (x86)\HP\BIOS Configuration Utility\BiosConfigUtility64.exe"
If fso.FileExists(filepath) then
Reporter.ReportEvent micPass,"BiosConfigUtility64.exe File Exists", ""
 var = "BiosConfigUtility64.exe File Exists"
 Call generatelogs (var)
 Else
Reporter.ReportEvent micFail, "BiosConfigUtility64.exe File doesnot Exists", ""
var = "BiosConfigUtility64.exe File doesnot Exists"
 Call generatelogs (var)
 End If
Set fso=nothing

Set fso=createobject("Scripting.FileSystemObject")
filepath="C:\Program Files (x86)\HP\BIOS Configuration Utility\HPQPswd.exe"
If fso.FileExists(filepath) then

Reporter.ReportEvent micPass, "HPQPswd.exe File Exists",""
var = "HPQPswd.exe File Exists"
 Call generatelogs (var)

 Else
Reporter.ReportEvent micFail, "HPQPswd.exe File doesnot Exist",""
var = "HPQPswd.exe File doesnot Exist"
 Call generatelogs (var)
End If
Set fso=nothing

Set fso=createobject("Scripting.FileSystemObject")
filepath="C:\Program Files (x86)\HP\BIOS Configuration Utility\HPQPswd64.exe"
If fso.FileExists(filepath) then
 Reporter.ReportEvent micPass,"HPQPswd64.exe File Exists", ""
 var = "HPQPswd64.exe File Exists"
 Call generatelogs (var)
Else
 Reporter.ReportEvent micPass,"HPQPswd64.exe File dosenot Exists", ""
  var = "HPQPswd64.exe File dosenot Exists"
 Call generatelogs (var)
End If
Set fso=nothing
'-----------Uninstall "HP BIOS Configuration Utility" through control panel----------------------------------------------

Call SW_Uninstallation ("HP BIOS Configuration Utility")


