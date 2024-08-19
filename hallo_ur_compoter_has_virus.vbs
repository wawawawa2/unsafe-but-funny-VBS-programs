' Set the script name
Dim scriptName
scriptName = "popup_spam.vbs"

' Create a Shell object
Set shell = CreateObject("WScript.Shell")

' Check if the script is already running
If shell.ExpandEnvironmentStrings("%RUNNING%") <> "" Then
    ' If the variable is set, show the popup
    MsgBox "Warning: This is a spam popup!", vbCritical, "Spam Alert"
    ' Exit the script
    WScript.Quit
End If

' Set an environment variable to indicate that the script is running
shell.Environment("Process")("RUNNING") = "1"

' Show a critical message box
MsgBox "Warning: This is a spam popup!", vbCritical, "Spam Alert"

' Run the script in a new process
shell.Run "wscript """ & WScript.ScriptFullName & """", 0, False

' Clean up
Set shell = Nothing
