
'Run as Admin command to the shell


Set WshShell = WScript.CreateObject("WScript.Shell")
If WScript.Arguments.Length = 0 Then
  Set ObjShell = CreateObject("Shell.Application")
  ObjShell.ShellExecute "wscript.exe" _
    , """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
  WScript.Quit
End if

'Make the shell stop superfecth, auto checks if it is disabled or not, return message is still the same


set aShell= CreateObject("Wscript.Shell")
aShell.run "%SYSTEMROOT%\system32\net.exe stop SysMain",0, True
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colServiceList = objWMIService.ExecQuery("Select * from Win32_Service " _
& "where Name='SysMain'")
For Each objService In colServiceList
StartMode = UCase(objService.StartMode)


If (StartMode = "MANUAL") Or (StartMode = "AUTO") Then
errReturnCode = objService.ChangeStartMode("Disabled")
End If
Next
Wscript.Echo "Super fetch has stopped and been disabled "

'I took out the :) because too many people thought I was running 1337 hax and kept killing my little program



'Added the ability to open IE, I know its IE but new machines don't have chrome enterprise instaled so dry your tears


Set objExplorer = CreateObject("InternetExplorer.Application")
WebSite = "https://www.ninite.com"
with objExplorer
.Navigate2 WebSite
.AddressBar = 1
.Visible = 1
.ToolBar = 1
.StatusBar = 1
end with


WebSite = "https://get.adobe.com/reader"
with objExplorer
.Navigate2 WebSite, &h800
.AddressBar = 1
.Visible = 1
.ToolBar = 1
.StatusBar = 1
end with


WebSite = "https://www.office.com"
with objExplorer
.Navigate2 WebSite, &h800
.AddressBar = 1
.Visible = 1
.ToolBar = 1
.StatusBar = 1
end with




'This will then open power options and the virtual memory boxes. 
'After you close memory box add + remove starts


set aShell= CreateObject("Wscript.Shell")
aShell.run "%windir%\system32\control.exe /name Microsoft.PowerOptions",1, True

set aShell= CreateObject("Wscript.Shell")
aShell.run "%windir%\system32\SystemPropertiesPerformance.exe",1, True

set aShell= CreateObject("Wscript.Shell")
aShell.run "%windir%\system32\control.exe /name Microsoft.ProgramsAndFeatures",1, True




