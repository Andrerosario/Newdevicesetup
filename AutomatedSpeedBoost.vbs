Set WshShell = WScript.CreateObject("WScript.Shell")
If WScript.Arguments.Length = 0 Then
  Set ObjShell = CreateObject("Shell.Application")
  ObjShell.ShellExecute "wscript.exe" _
    , """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
  WScript.Quit
End if

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
Wscript.Echo "Super fetch has stopped and been disabled :) "





set aShell= CreateObject("Wscript.Shell")
aShell.run "%windir%\system32\control.exe /name Microsoft.PowerOptions",1, True

set aShell= CreateObject("Wscript.Shell")
aShell.run "%windir%\system32\SystemPropertiesPerformance.exe",1, True

set aShell= CreateObject("Wscript.Shell")
aShell.run "%windir%\system32\control.exe /name Microsoft.ProgramsAndFeatures",1, True







