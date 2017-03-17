'*********************************************************************************************************************************************
'Name                 :  Jagadish           
'Description          :  Step 001 - BCU softpaq extraction
'Pre-Requisite        :  C:\CMIT\ folder need to create 2.Setup exe file should place
'Created By           :      
'Creation Date        :                     
'Application Name     :                    
'Changed By         Date @ Time:            Description:                        
'**********************************************************************************************************************************************

'========Prerequsite 1.C:\CMIT\ folder need to create 2.Setup exe file should place=============
systemutil.Run "C:\CMIT\Setup.exe"
wait 2
Window("InstallShield Wizard").WinButton("Next >").Click
Window("InstallShield Wizard").WinButton("Next >").Click
Window("InstallShield Wizard").WinButton("Install").Click
Window("InstallShield Wizard").WinButton("Finish").Click

'-------------verify installation done sucessfully-----------
Set WshShell = CreateObject("WScript.Shell")
WshShell.SendKeys "(^{ESC}{TAB})"
Wait 2
WshShell.SendKeys "appwiz.cpl"
Wait 2
WshShell.SendKeys "{Enter}"
wait 2
Window("Programs and Features").WinObject("Search Programs and Features").WinEdit("Search Box").Set "HP BIOS Configuration Utility"
Set WshShell = CreateObject("wscript.Shell")
wait 3
WshShell.SendKeys "{down}"
Window("Programs and Features").WinListView("WinListView").Click
WshShell.SendKeys "{down}"


var=Window("Programs and Features").WinListView("WinListView").GetROProperty ("Selection")
 if var= "HP BIOS Configuration Utility" then 
Call generatelogs (var)
Else 
var = "HP BIOS Configuration Utility not exists"
Call generatelogs (var)
End if 

Window("Programs and Features").Close

WshShell.SendKeys "Exit"
wait 2
WshShell.SendKeys "{Enter}"
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



