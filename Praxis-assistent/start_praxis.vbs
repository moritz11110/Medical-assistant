Option Explicit

Dim basis, pythonExe, pythonwExe, mainCmd
Dim shell, fso

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

basis = fso.GetParentFolderName(WScript.ScriptFullName)
pythonExe = basis & "\\.venv\\Scripts\\python.exe"
pythonwExe = basis & "\\.venv\\Scripts\\pythonw.exe"

If Not fso.FileExists(pythonExe) Then
    pythonExe = "python"
End If

If Not fso.FileExists(pythonwExe) Then
    pythonwExe = "pythonw"
End If

mainCmd = Chr(34) & pythonwExe & Chr(34) & " " & Chr(34) & basis & "\\main.py" & Chr(34)

' App versteckt starten, nicht wartend
Call shell.Run(mainCmd, 0, False)
