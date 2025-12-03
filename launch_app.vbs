' -------------------------------
' Universal Launcher for Streamlit App
' -------------------------------

Option Explicit
Dim WshShell, AppPath, PythonExe, Command

' Get the folder where this VBS is located
AppPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

Set WshShell = WScript.CreateObject("WScript.Shell")

' Check if virtual environment exists
If WshShell.CurrentDirectory <> AppPath Then
    WshShell.CurrentDirectory = AppPath
End If

' Define the Python executable
' If venv exists, use venv\Scripts\python.exe
If CreateObject("Scripting.FileSystemObject").FolderExists(AppPath & "\venv") Then
    PythonExe = AppPath & "\venv\Scripts\python.exe"
Else
    ' If no venv, use default python in PATH
    PythonExe = "python"
End If

' Command to run Streamlit app
Command = PythonExe & " -m streamlit run """ & AppPath & "\app.py"""

' Run the command in a new window
WshShell.Run Command, 1, False
