Option Explicit

Dim strProcessName, WshShell, Excel
Set Excel = WScript.CreateObject("Excel.Application")
strProcessName = "wscript.exe"
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
    'stop space key
    Const VK_SPACE = &H20 ' The SPACEBAR key.
    Call KeybordEvent(VK_SPACE, 0, 2, 0)
    Call KeybordEvent(VK_SPACE, 0, 2, 0)
    WScript.Sleep 60000 ' 60000 milliseconds equals 1 minutes
    Call Start() ' Move the mouse again
End Sub


Sub MoveMouse()
    Dim WshShell, intScreenWidth, intScreenHeight, command
    Randomize ' Initialize the random number seed
    ' Get the screen size
    intScreenWidth = 1000 ' Get the screen width
    intScreenHeight = 1000 ' Get the screen height
    ' Generate random new position
    Dim newX, newY
    newX = Int(Rnd * intScreenWidth) ' Generate a random X coordinate
    newY = Int(Rnd * intScreenHeight) ' Generate a random Y coordinate
    command = "CALL(""user32.dll"", ""SetCursorPos"", ""JJJ"", "& newX &", "& newY &")"
    Excel.ExecuteExcel4Macro(command)
    ' Start the timer, move the mouse again after 3 minutes
End Sub

Dim Minus
Minus = True
Sub MouseWheelEvent()
    Const MOUSEEVENTF_WHEEL = &H800 ' The wheel was rolled.
    Randomize ' Initialize the random number seed
    Dim randNum
    If Minus Then
        randNum = Int(Rnd * 300) ' Generate a random number
        Minus = False
    Else
        randNum = Int(Rnd * 300) * -1 ' Generate a random number
        Minus = True
    End If
    Call MouseEvent(MOUSEEVENTF_WHEEL, 0, 0, randNum, 0)
End sub


Public Sub MouseClick()
    Const MOUSEEVENTF_LEFTDOWN = &H2 ' The left button was pressed.
    Const MOUSEEVENTF_LEFTUP = &H4 ' The left button was released.
    Const VK_CTL = &H11 ' The CTRL key.
    Const VK_SHIFT = &H10 ' The SHIFT key.
    Const VK_ENTER = &HD ' The ENTER key.
    Const VK_SPACE = &H20 ' The SPACEBAR key.
    ' Key Release
    Call KeybordEvent(VK_SPACE, 0, 2, 0)
    Dim dwFlags
    dwFlags = MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP
    Call KeybordEvent(VK_CTL, 0, 3, 0)
    Call MouseEvent(dwFlags, 0, 0, 0, 0)
End Sub


Sub MouseEvent(dwFlags, dx, dy, dwData, dwExtraInfo)
    Dim strFunction
    Const command = "CALL(""user32"",""mouse_event"",""JJJJJj"", $1, $2, $3, $4, $5)"
    strFunction = Replace(Replace(Replace(Replace(Replace(command, "$1", dwFlags), "$2", dx), "$3", dy), "$4", dwData), "$5", dwExtraInfo)
    Call Excel.ExecuteExcel4Macro(strFunction)
End Sub


Sub KeybordEvent(bVk, bScan, dwFlags, dwExtraInfo)
    Dim strFunction
    Const command = "CALL(""user32"",""keybd_event"",""JJJJJ"", $1, $2, $3, $4)"
    strFunction = Replace(Replace(Replace(Replace(command, "$1", bVk), "$2", bScan), "$3", dwFlags), "$4", dwExtraInfo)
    Call Excel.ExecuteExcel4Macro(strFunction)
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
