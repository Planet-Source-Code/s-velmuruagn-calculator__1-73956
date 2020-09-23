Attribute VB_Name = "Module1"
Global Const HELP_CONTEXT = &H1
Global Const HELP_QUIT = &H2
Global Const HELP_FINDER = &HB
Global Const HELP_INDEX = &H3
Global Const HELP_HELPONHELP = &H4
Global Const HELP_SETINDEX = &H5
Global Const HELP_KEY = &H101
Global Const HELP_MULTIKEY = &H201
Global Const HELP_CONTENTS = &H3
Global Const HELP_SETCONTENTS = &H5
Global Const HELP_CONTEXTPOPUP = &H8
Global Const HELP_FORCEFILE = &H9
Global Const HELP_COMMAND = &H102
Global Const HELP_PARTIALKEY = &H105
Global Const HELP_SETWINPOS = &H203

    Declare Function WinHelpByNum Lib "User32.dll" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData&) As Long
    Dim m_hWndMainWindow As Long
Public Sub ShowHelpContents()
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTENTS, CLng(0))

End Sub

