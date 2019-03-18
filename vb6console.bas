Attribute VB_Name = "vb6console"
Option Explicit

'For a new console project:
'Add reference to Microsoft Scripting Runtime via Project->References
'Link to console subsytem after compile to direct output to executing console vs. new console
'Your location of link.exe may vary, but for example:
'"C:\Program Files (x86)\Microsoft Visual Studio\vb98\LINK.EXE" /EDIT /SUBSYSTEM:CONSOLE <exename>
'NOTE: you will have to re-link after each compile
'
'Base code is an edit from https://stackoverflow.com/questions/10517338/how-to-write-to-a-debug-console-in-vb6/10517370

Public SIn As Scripting.TextStream
Public SOut As Scripting.TextStream

Private Sub Main()

    With New Scripting.FileSystemObject
        Set SIn = .GetStandardStream(StdIn)
        Set SOut = .GetStandardStream(StdOut)
    End With

    SOut.WriteLine "Any output you want"
    SOut.WriteLine "Goes here"

End Sub

