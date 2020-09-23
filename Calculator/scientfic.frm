VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scientific Calculator"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8490
   HelpContextID   =   4
   Icon            =   "scientfic.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8490
   Begin VB.CommandButton Command28 
      BackColor       =   &H00FF8080&
      Caption         =   "Round"
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H00FF80FF&
      Caption         =   "Del"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command26 
      BackColor       =   &H00FF8080&
      Caption         =   "Hex"
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H00FF8080&
      Caption         =   "Binary"
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00FF8080&
      Caption         =   "Oct"
      Height          =   495
      Left            =   7200
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00FF8080&
      Caption         =   "Not"
      Height          =   495
      Left            =   6360
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00FF8080&
      Caption         =   "And"
      Height          =   495
      Left            =   6360
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00FF8080&
      Caption         =   "Or"
      Height          =   495
      Left            =   7200
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00FF8080&
      Caption         =   "log"
      Height          =   495
      Left            =   5520
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00FF8080&
      Caption         =   "x!"
      Height          =   495
      Left            =   6360
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00FF8080&
      Caption         =   "x^3"
      Height          =   495
      Left            =   6360
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FF8080&
      Caption         =   "Int"
      Height          =   495
      Left            =   6360
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00FF8080&
      Caption         =   "Tan"
      Height          =   495
      Left            =   5520
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00FF8080&
      Caption         =   "Cos"
      Height          =   495
      Left            =   5520
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00FF8080&
      Caption         =   "Sin"
      Height          =   495
      Left            =   5520
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00FF8080&
      Caption         =   "Pi"
      Height          =   495
      Left            =   5520
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "1"
      Height          =   495
      Index           =   0
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "2"
      Height          =   495
      Index           =   1
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "3"
      Height          =   495
      Index           =   2
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "4"
      Height          =   495
      Index           =   3
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "5"
      Height          =   495
      Index           =   4
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "6"
      Height          =   495
      Index           =   5
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "7"
      Height          =   495
      Index           =   6
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "8"
      Height          =   495
      Index           =   7
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "9"
      Height          =   495
      Index           =   8
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "0"
      Height          =   495
      Index           =   9
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4200
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
      Left            =   2760
      TabIndex        =   17
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00008000&
      Caption         =   "+/-"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF80FF&
      Caption         =   "."
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton cmd_mem_clear 
      BackColor       =   &H00FF80FF&
      Caption         =   "MC"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmd_mem_recol 
      BackColor       =   &H00FF80FF&
      Caption         =   "MR"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmd_ms 
      BackColor       =   &H00FF80FF&
      Caption         =   "MS"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF80FF&
      Caption         =   "M+"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF80FF&
      Caption         =   "<-backSpace"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00008000&
      Caption         =   "+"
      Height          =   495
      Index           =   0
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00008000&
      Caption         =   "-"
      Height          =   495
      Index           =   1
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00008000&
      Caption         =   "*"
      Height          =   495
      Index           =   2
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00008000&
      Caption         =   "/"
      Height          =   495
      Index           =   3
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FF80FF&
      Caption         =   "CE"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FF80FF&
      Caption         =   "C"
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00008000&
      Caption         =   "Sqrt"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00008000&
      Caption         =   "1/x"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FF80FF&
      Caption         =   "="
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H0000C000&
      Caption         =   "x*x"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BorderColor     =   &H000000FF&
      Height          =   3855
      Left            =   5280
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H000080FF&
      Height          =   3855
      Left            =   120
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label lbl_memory 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   360
      TabIndex        =   28
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
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
Call Command13_KeyPress(KeyAscii)
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
Call Command13_KeyPress(KeyAscii)
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
Call Command13_KeyPress(KeyAscii)
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
Call Command13_KeyPress(KeyAscii)
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
Call Command13_KeyPress(KeyAscii)
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
If operator = "OR" Then Text1.Text = operend Or val(Text1.Text)
If operator = "AND" Then Text1.Text = operend And val(Text1.Text)
End Sub

Private Sub Command11_GotFocus()
Command11.BackColor = &HFF8080
End Sub

Private Sub Command11_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
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
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command12_LostFocus()
Command12.BackColor = &H8000&
End Sub

Private Sub Command13_Click()
Text1.Text = 22 / 7
End Sub

Private Sub Command13_GotFocus()
Command13.BackColor = &HFFFF80 '&H00FF8080&
End Sub

Private Sub Command13_KeyPress(KeyAscii As Integer)
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

Private Sub Command13_LostFocus()
Command13.BackColor = &HFF8080
End Sub

Private Sub Command14_Click()
On Error Resume Next
Text1.Text = Sin(val(Text1.Text))
End Sub

Private Sub Command14_GotFocus()
Command14.BackColor = &HFFFF80
End Sub

Private Sub Command14_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command14_LostFocus()
Command14.BackColor = &HFF8080
End Sub

Private Sub Command15_Click()
On Error Resume Next
Text1.Text = Cos(val(Text1.Text))
End Sub

Private Sub Command15_GotFocus()
Command15.BackColor = &HFFFF80
End Sub

Private Sub Command15_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command15_LostFocus()
Command15.BackColor = &HFF8080
End Sub

Private Sub Command16_Click()
On Error Resume Next
Text1.Text = Tan(val(Text1.Text))
End Sub

Private Sub Command16_GotFocus()
Command16.BackColor = &HFFFF80
End Sub

Private Sub Command16_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command16_LostFocus()
Command16.BackColor = &HFF8080
End Sub

Private Sub Command17_Click()
Text1.Text = Int(val(Text1.Text))
End Sub

Private Sub Command17_GotFocus()
Command17.BackColor = &HFFFF80
End Sub

Private Sub Command17_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command17_LostFocus()
Command17.BackColor = &HFF8080
End Sub

Private Sub Command18_Click()
Text1.Text = val(Text1.Text) * val(Text1.Text) * val(Text1.Text)
End Sub

Private Sub Command18_GotFocus()
Command18.BackColor = &HFFFF80
End Sub

Private Sub Command18_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command18_LostFocus()
Command18.BackColor = &HFF8080
End Sub

Private Sub Command19_Click()
Dim i As Double
Dim val As Double
val = 1
For i = 1 To Text1.Text
val = val * i
Next
Text1.Text = val

End Sub

Private Sub Command19_GotFocus()
Command19.BackColor = &HFFFF80
End Sub

Private Sub Command19_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command19_LostFocus()
Command19.BackColor = &HFF8080
End Sub

Private Sub Command2_Click()
Text1.Text = val(Text1.Text) * -1
End Sub

Private Sub Command2_GotFocus()
Command2.BackColor = &HC000C0
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command2_LostFocus()
Command2.BackColor = &H8000&
End Sub

Private Sub Command20_Click()
Text1.Text = Log(val(Text1.Text))
End Sub

Private Sub Command20_GotFocus()
Command20.BackColor = &HFFFF80
End Sub

Private Sub Command20_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command20_LostFocus()
Command20.BackColor = &HFF8080
End Sub

Private Sub Command21_Click()
operend = val(Text1.Text)
operator = "OR"
Text1.Text = 0
'Text1.Text = operend Or val(Text1.Text)
End Sub

Private Sub Command21_GotFocus()
Command21.BackColor = &HFFFF80
End Sub

Private Sub Command21_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command21_LostFocus()
Command21.BackColor = &HFF8080
End Sub

Private Sub Command22_Click()

operend = val(Text1.Text)
operator = "AND"
Text1.Text = 0
End Sub

Private Sub Command22_GotFocus()
Command22.BackColor = &HFFFF80
End Sub

Private Sub Command22_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command22_LostFocus()
Command22.BackColor = &HFF8080
End Sub

Private Sub Command23_Click()
Text1.Text = Not (val(Text1.Text))
End Sub

Private Sub Command23_GotFocus()
Command23.BackColor = &HFFFF80
End Sub

Private Sub Command23_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command23_LostFocus()
Command23.BackColor = &HFF8080
End Sub

Private Sub Command24_Click()
Text1.Text = Oct(val(Text1.Text))
End Sub

Private Sub Command24_GotFocus()
Command24.BackColor = &HFFFF80
End Sub

Private Sub Command24_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command24_LostFocus()
Command24.BackColor = &HFF8080
End Sub

Private Sub Command25_Click()
Dim df As Integer
Dim r As Integer
df = val(Text1.Text)
Text1.Text = ""
Do While (df)
r = MDIForm1.reminder(df, 2)
Text1.Text = r & Text1.Text
df = Int(df / 2)
Loop
End Sub
Private Sub Command25_GotFocus()
Command25.BackColor = &HFFFF80
End Sub

Private Sub Command25_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command25_LostFocus()
Command25.BackColor = &HFF8080
End Sub

Private Sub Command26_Click()
Text1.Text = Hex$(val(Text1.Text))
End Sub
Private Sub Command26_GotFocus()
Command26.BackColor = &HFFFF80
End Sub

Private Sub Command26_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command26_LostFocus()
Command26.BackColor = &HFF8080
End Sub

Private Sub Command27_Click()
Text1.Text = Mid(Text1.Text, 2, Len(Text1.Text))
End Sub
Private Sub Command27_GotFocus()
Command27.BackColor = &HFF8080
End Sub

Private Sub Command27_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command27_LostFocus()
Command27.BackColor = &HFF80FF
End Sub

Private Sub Command28_Click()
On Error Resume Next
a = Int(Abs(InputBox("Enter Number of decimal places ", "Scientifc Calculator", 0)))

Text1.Text = Round(val(Text1.Text), a)
End Sub
Private Sub Command28_GotFocus()
Command28.BackColor = &HFFFF80
End Sub

Private Sub Command28_KeyPress(KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command28_LostFocus()
Command28.BackColor = &HFF8080
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
Call Command13_KeyPress(KeyAscii)
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
Call Command13_KeyPress(KeyAscii)
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
Call Command13_KeyPress(KeyAscii)
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
If operator = "OR" Then Text1.Text = operend Or val(Text1.Text)
If operator = "AND" Then Text1.Text = operend And val(Text1.Text)

End If
 operator = Command6(Index).Caption
 Text1.Text = 0
End Sub

Private Sub Command6_GotFocus(Index As Integer)
Command6(Index).BackColor = &HC000C0
End Sub

Private Sub Command6_KeyPress(Index As Integer, KeyAscii As Integer)
Call Command13_KeyPress(KeyAscii)
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
Call Command13_KeyPress(KeyAscii)
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
Call Command13_KeyPress(KeyAscii)
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
Call Command13_KeyPress(KeyAscii)
End Sub

Private Sub Command9_LostFocus()
Command9.BackColor = &H8000&
End Sub

Private Sub Form_Load()
co = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.scientific.Checked = False
End Sub

Private Sub Text1_Change()
Text1.Locked = True
End Sub

Private Sub Timer1_Timer()
co = co + 1
Shape2.BorderColor = QBColor(MDIForm1.reminder(co, 15))
If co = 16 Then co = 1
End Sub


