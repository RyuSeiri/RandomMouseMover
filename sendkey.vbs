Option Explicit
Dim strProcessName, WshShell, Excel
strProcessName = "wscript.exe"
Set WshShell = WScript.CreateObject("WScript.Shell")

' Check if the script is already running
If IsScriptRunning()  Then
    Call CloseScript()
Else
    ' If the script is not running, start the script and display info
    MsgBox "started", vbInformation, "info"
    Call Start()
End If


Function IsScriptRunning()
    Dim objWMIService, colProcess, objProcess, processCount
    processCount = 0
    ' Get the running script process
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & "." & "\root\cimv2")
    Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & strProcessName & "'")
    
    ' Check if the script is already running
    For Each objProcess In colProcess
        If objProcess.Name = strProcessName Then
            processCount = processCount + 1
        End If
    Next
    If processCount > 1 Then
        IsScriptRunning = True
    Else
        IsScriptRunning = False
    End If
End Function


Sub Start()
    ' Send the Enter key
    WshShell.SendKeys "{ENTER}"
    WScript.Sleep 10000
    Call Start()
End Sub


Sub CloseScript()
    Dim objWMIService, colProcess, objProcess
    ' Get the running script process
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & "." & "\root\cimv2")
    Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & strProcessName & "'")

    ' Close the script and display the closing message
    MsgBox "Stopped", vbInformation, "info"
    For Each objProcess In colProcess
        If objProcess.Name = strProcessName Then
            objProcess.Terminate()
        End If
    Next
End Sub