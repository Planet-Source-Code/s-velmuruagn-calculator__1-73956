VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Calculator"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5490
   HelpContextID   =   3
   Icon            =   "simple calculator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   5490
   Begin VB.CommandButton Command27 
      BackColor       =   &H00FF80FF&
      Caption         =   "Del"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1080
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   4920
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H0000C000&
      Caption         =   "x*x"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FF80FF&
      Caption         =   "="
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00008000&
      Caption         =   "1/x"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00008000&
      Caption         =   "Sqrt"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FF80FF&
      Caption         =   "C"
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FF80FF&
      Caption         =   "CE"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00008000&
      Caption         =   "/"
      Height          =   495
      Index           =   3
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00008000&
      Caption         =   "*"
      Height          =   495
      Index           =   2
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00008000&
      Caption         =   "-"
      Height          =   495
      Index           =   1
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00008000&
      Caption         =   "+"
      Height          =   495
      Index           =   0
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF80FF&
      Caption         =   "<-backSpace"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF80FF&
      Caption         =   "M+"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmd_ms 
      BackColor       =   &H00FF80FF&
      Caption         =   "MS"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmd_mem_recol 
      BackColor       =   &H00FF80FF&
      Caption         =   "MR"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmd_mem_clear 
      BackColor       =   &H00FF80FF&
      Caption         =   "MC"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF80FF&
      Caption         =   "."
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00008000&
      Caption         =   "+/-"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.000E+00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   6
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "0"
      Height          =   495
      Index           =   9
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "9"
      Height          =   495
      Index           =   8
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "8"
      Height          =   495
      Index           =   7
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "7"
      Height          =   495
      Index           =   6
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "6"
      Height          =   495
      Index           =   5
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "5"
      Height          =   495
      Index           =   4
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "4"
      Height          =   495
      Index           =   3
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "3"
      Height          =   495
      Index           =   2
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "2"
      Height          =   495
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "1"
      Height          =   495
      Index           =   0
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C000&
      Height          =   615
      Left            =   240
      Top             =   240
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   3735
      Left            =   240
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label lbl_memory 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FF80&
      FillColor       =   &H00FF0000&
      Height          =   4695
      Left            =   120
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim number As Double
Dim operend As Double
Dim operator As String
Dim co As Integer
Private Sub cmd_mem_clear_Click()
lbl_memory.Caption = ""
number = 0
End Sub

Private Sub cmd_mem_clear_GotFocus()
cmd_mem_clear.BackColor = &HFF8080 '&H00FF80FF&
End Sub

Private Sub cmd_mem_clear_KeyPress(KeyAscii As Integer)
Call Command11_KeyPress(KeyAscii)
End Sub

Private Sub cmd_mem_clear_LostFocus()
cmd_mem_clear.BackColor = &HFF80FF
End Sub

Private Sub cmd_mem_recol_Click()
Text1.Text = number
End Sub

Private Sub cmd_mem_recol_GotFocus()
cmd_mem_recol.BackColor = &HFF8080
End Sub

Private Sub cmd_mem_recol_KeyPress(KeyAscii As Integer)
Call Command11_KeyPress(KeyAscii)
End Sub

Private Sub cmd_mem_recol_LostFocus()
cmd_mem_recol.BackColor = &HFF80FF
End Sub

Private Sub cmd_ms_Click()

number = val(Text1.Text)
lbl_memory.Caption = "M"
End Sub

Private Sub cmd_ms_GotFocus()
cmd_ms.BackColor = &HFF8080
End Sub

Private Sub cmd_ms_KeyPress(KeyAscii As Integer)
Call Command11_KeyPress(KeyAscii)
End Sub

Private Sub cmd_ms_LostFocus()
cmd_ms.BackColor = &HFF80FF
End Sub

Private Sub Command1_Click(Index As Integer)
For i = 1 To Len(Text1.Text)
If Mid(Text1.Text, i, 1) >= "A" Then
Text1.Text = 0
Exit For
End If
Next i

Text1.Text = Text1.Text & Command1(Index).Caption
End Sub

Private Sub Command1_GotFocus(Index As Integer)
Command1(Index).BackColor = &HC0FFC0
End Sub

Private Sub Command1_KeyPress(Index As Integer, KeyAscii As Integer)
Call Command11_KeyPress(KeyAscii)
End Sub

Private Sub Command1_LostFocus(Index As Integer)

Command1(Index).BackColor = &H80FF& '&H00FF8080&
End Sub

Private Sub Command10_Click()
If val(Text1.Text) <> 0 Then Text1.Text = 1 / val(Text1.Text)
End Sub

Private Sub Command10_GotFocus()
Command10.BackColor = &HC000C0
End Sub

Private Sub Command10_KeyPress(KeyAscii As Integer)
Call Command11_KeyPress(KeyAscii)
End Sub

Private Sub Command10_LostFocus()
Command10.BackColor = &H8000&
End Sub

Private Sub Command11_Click()

If operator = "+" Then Text1.Text = operend + val(Text1.Text)
If operator = "-" Then Text1.Text = operend - val(Text1.Text)
If operator = "*" Then Text1.Text = operend * val(Text1.Text)
If operator = "/" And val(Text1.Text) <> 0 Then Text1.Text = operend / val(Text1.Text)
If operator = "/" And val(Text1.Text) = 0 Then Text1.Text = "Cannot divide by Zero"
End Sub

Private Sub Command11_GotFocus()
Command11.BackColor = &HFF8080
End Sub

Private Sub Command11_KeyPress(KeyAscii As Integer)
If (Chr(KeyAscii) < Chr(Asc("0")) Or Chr(KeyAscii) > Chr(Asc("9"))) And Chr(KeyAscii) <> Chr(Asc(".")) Then
KeyAscii = 0
Exit Sub
End If
If Chr(KeyAscii) = Chr(Asc(".")) Then
For i = 1 To Len(Text1.Text)
If Mid(Text1.Text, i, 1) = "." Then
Exit Sub
End If
Next i
End If
Text1.Text = Text1.Text & Chr(KeyAscii)
End Sub

Private Sub Command11_LostFocus()
Command11.BackColor = &HFF80FF
End Sub

Private Sub Command12_Click()
Text1.Text = val(Text1.Text) * val(Text1.Text)
End Sub

Private Sub Command12_GotFocus()
Command12.BackColor = &HC000C0
End Sub

Private Sub Command12_KeyPress(KeyAscii As Integer)
Call Command11_KeyPress(KeyAscii)
End Sub

Private Sub Command12_LostFocus()
Command12.BackColor = &H8000&
End Sub

Private Sub Command2_Click()
Text1.Text = val(Text1.Text) * -1
End Sub

Private Sub Command2_GotFocus()
Command2.BackColor = &HC000C0
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
Call Command11_KeyPress(KeyAscii)
End Sub

Private Sub Command2_LostFocus()
Command2.BackColor = &H8000&
End Sub

Private Sub Command27_Click()
Text1.Text = Mid(Text1.Text, 2, Len(Text1.Text))
End Sub
Private Sub Command27_GotFocus()
Command27.BackColor = &HFF8080
End Sub

Private Sub Command27_KeyPress(KeyAscii As Integer)
Call Command11_KeyPress(KeyAscii)
End Sub

Private Sub Command27_LostFocus()
Command27.BackColor = &HFF80FF
End Sub

Private Sub Command3_Click()
For i = 1 To Len(Text1.Text)
If Mid(Text1.Text, i, 1) = "." Then
Exit Sub
End If
Next i
Text1.Text = Text1.Text & "."
End Sub

Private Sub Command3_GotFocus()
Command3.BackColor = &HFF8080
End Sub

Private Sub Command3_KeyPress(KeyAscii As Integer)
Call Command11_KeyPress(KeyAscii)
End Sub

Private Sub Command3_LostFocus()
Command3.BackColor = &HFF80FF
End Sub

Private Sub Command4_Click()
number = val(Text1.Text) + val(Text1.Text)
lbl_memory.Caption = "M"
End Sub

Private Sub Command4_GotFocus()
Command4.BackColor = &HFF8080
End Sub

Private Sub Command4_KeyPress(KeyAscii As Integer)
Call Command11_KeyPress(KeyAscii)
End Sub

Private Sub Command4_LostFocus()
Command4.BackColor = &HFF80FF
End Sub

Private Sub Command5_Click()
If Len(Text1.Text) <> 0 Then
Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 1)
End If
End Sub

Private Sub Command5_GotFocus()

Command5.BackColor = &HFF8080
End Sub

Private Sub Command5_KeyPress(KeyAscii As Integer)
Call Command11_KeyPress(KeyAscii)
End Sub

Private Sub Command5_LostFocus()
Command5.BackColor = &HFF80FF
End Sub

Private Sub Command6_Click(Index As Integer)
If operend = 0 Then
operend = operend + val(Text1.Text)
Else
If operator = "+" Then operend = operend + val(Text1.Text)
If operator = "-" Then operend = operend - val(Text1.Text)
If operator = "*" Then operend = operend * val(Text1.Text)
If operator = "/" And val(Text1.Text) <> 0 Then operend = operend / val(Text1.Text)
If operator = "/" And val(Text1.Text) = 0 Then Text1.Text = "Cannot divide by Zero": Exit Sub

End If
 operator = Command6(Index).Caption
 Text1.Text = 0
End Sub

Private Sub Command6_GotFocus(Index As Integer)
Command6(Index).BackColor = &HC000C0
End Sub

Private Sub Command6_KeyPress(Index As Integer, KeyAscii As Integer)
Call Command11_KeyPress(KeyAscii)
End Sub

Private Sub Command6_LostFocus(Index As Integer)
Command6(Index).BackColor = &H8000&
End Sub

Private Sub Command7_Click()
Text1.Text = ""
End Sub

Private Sub Command7_GotFocus()
Command7.BackColor = &HFF8080
End Sub

Private Sub Command7_KeyPress(KeyAscii As Integer)
Call Command11_KeyPress(KeyAscii)
End Sub

Private Sub Command7_LostFocus()
Command7.BackColor = &HFF80FF
End Sub

Private Sub Command8_Click()
Text1.Text = ""
operend = 0
End Sub

Private Sub Command8_GotFocus()
Command8.BackColor = &HFF8080
End Sub

Private Sub Command8_KeyPress(KeyAscii As Integer)
Call Command11_KeyPress(KeyAscii)
End Sub

Private Sub Command8_LostFocus()
Command8.BackColor = &HFF80FF
End Sub

Private Sub Command9_Click()
If Mid(Text1.Text, 1, 1) <> "-" Then
Text1.Text = Sqr(val(Text1.Text))
Else
Text1.Text = "Invalid input for function"
End If
End Sub

Private Sub Command9_GotFocus()
Command9.BackColor = &HC000C0
End Sub

Private Sub Command9_KeyPress(KeyAscii As Integer)
Call Command11_KeyPress(KeyAscii)
End Sub

Private Sub Command9_LostFocus()
Command9.BackColor = &H8000&
End Sub

Private Sub Form_Load()
co = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Standard.Checked = False
End Sub

Private Sub Text1_Change()
Text1.Locked = True
End Sub

Private Sub Timer1_Timer()
co = co + 1
Shape2.BorderColor = QBColor(MDIForm1.reminder(co, 15))
If co = 16 Then co = 1
End Sub

