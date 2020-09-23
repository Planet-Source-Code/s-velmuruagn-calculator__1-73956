VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00404040&
   Caption         =   "Main Form"
   ClientHeight    =   5625
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6180
   HelpContextID   =   2
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu paste 
         Caption         =   "Paste"
      End
      Begin VB.Menu clear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu view 
      Caption         =   "View"
      Begin VB.Menu Standard 
         Caption         =   "Standard"
      End
      Begin VB.Menu scientific 
         Caption         =   "Scientific"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu helptopic 
         Caption         =   "Help Topic"
      End
      Begin VB.Menu sp 
         Caption         =   "-"
      End
      Begin VB.Menu aboutus 
         Caption         =   "About us"
      End
      Begin VB.Menu aboutcalculator 
         Caption         =   "About Calculator"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double

Private Sub aboutcalculator_Click()
frmAbout.Show
End Sub

Private Sub aboutus_Click()
Form3.Show
End Sub

Private Sub clear_Click()
On Error Resume Next
Me.ActiveForm.Text1.Text = ""
End Sub

Private Sub copy_Click()
On Error Resume Next
a = Me.ActiveForm.Text1.Text
End Sub

Private Sub helptopic_Click()
ShowHelpContents
End Sub

Private Sub MDIForm_Load()
a = 0
Form1.Show
Standard.Checked = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub paste_Click()
On Error Resume Next
 Me.ActiveForm.Text1.Text = a
End Sub

Private Sub scientific_Click()
Unload Form1
Form2.Show
Standard.Checked = False
scientific.Checked = True
End Sub

Private Sub Standard_Click()
Unload Form2
Form1.Show
Standard.Checked = True
scientific.Checked = False
End Sub

Public Function reminder(vl As Integer, divi As Integer) As Integer
Dim qua As Integer
If divi <> 0 Then
qua = Round(vl \ divi, 0)
reminder = vl - (qua * divi)
Else
reminder = 0
End If
End Function
