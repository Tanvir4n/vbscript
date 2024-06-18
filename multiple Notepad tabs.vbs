' This VBScript example opens multiple instances of Notepad and arranges them randomly on the screen:
Set shell = WScript.CreateObject("WScript.Shell")

For i = 1 To 10
    shell.Run "notepad.exe"
    WScript.Sleep 100
Next

WScript.Sleep 1000

Set wmi = GetObject("winmgmts:").InstancesOf("Win32_Process")
For Each process in wmi
    If process.Name = "notepad.exe" Then
        ' Get the process ID and find the window handle
        Set processItem = GetObject("winmgmts:Win32_Process.Handle='" & process.ProcessId & "'")
        hwnd = shell.AppActivate(processItem.ProcessId)
        
        ' Move and resize the window randomly
        shell.SendKeys "%{ }"
        WScript.Sleep 200
    End If
Next
