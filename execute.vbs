Option Explicit

Dim strProcessName, WshShell
strProcessName = "wscript.exe"
' Check if the script is already running
If IsScriptRunning()  Then
    Call CloseScript()
Else
    ' If the script is not running, start the script and display info
    MsgBox "started", vbInformation, "info"
    Call MoveMouse()
End If


Sub MoveMouse()
    Dim WshShell, intScreenWidth, intScreenHeight, Excel, Command
    Randomize ' Initialize the random number seed
    ' Get the screen size
    intScreenWidth = 1000 ' Get the screen width
    intScreenHeight = 1000 ' Get the screen height
    ' Generate random new position
    Dim newX, newY
    newX = Int(Rnd * intScreenWidth) ' Generate a random X coordinate
    newY = Int(Rnd * intScreenHeight) ' Generate a random Y coordinate
    Set Excel = WScript.CreateObject("Excel.Application")
    Command = "CALL(""user32.dll"", ""SetCursorPos"", ""JJJ"", "& newX &", "& newY &")"
    Excel.ExecuteExcel4Macro(command)
    ' Start the timer, move the mouse again after 3 minutes
    WScript.Sleep 180000 ' 180000 milliseconds equals 3 minutes
    Call MoveMouse() ' Move the mouse again
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
