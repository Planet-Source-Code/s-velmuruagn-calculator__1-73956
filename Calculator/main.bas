Attribute VB_Name = "Module2"
Public Sub Main()
On Error GoTo vel_error
Dim strcmdline As String
strcmdline = Left(Command, 2)
If App.PrevInstance = True Then
MsgBox "Allready Program is running", vbOKOnly + vbInformation + vbMsgBoxHelpButton, "Calculator", App.HelpFile, 1
Exit Sub
End If
Select Case strcmdline
Case "\h"
ShowHelpContents
Case "\H"
ShowHelpContents
Case "\s"
MsgBox "Click help to open the command line help", vbMsgBoxHelpButton, "Calculator", App.HelpFile, 7
Case Else
MDIForm1.Show
End Select
Exit_main:
Exit Sub
vel_error:
#If gndebug Then
Stop
Resume
#End If
handleerror "main", Err.Description, Err.number, gErrFormName
Resume Exit_main
End Sub

Public Sub handleerror(loc As String, strerror$, lerror As Long, varmodule As Variant)
Dim ourcursortype As Integer
ourcursortype = Screen.MousePointer
Screen.MousePointer = vbNormal
MsgBox loc & ":" & strerror & "(" & lerror & ")", vbExclamation, varmodule + " Error"
Screen.MousePointer = ourcursortype
End Sub

