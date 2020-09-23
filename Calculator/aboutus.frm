VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About us"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3960
   HelpContextID   =   6
   Icon            =   "aboutus.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   3960
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3000
      Top             =   1560
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   $"aboutus.frx":0442
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   1080
      Picture         =   "aboutus.frx":04ED
      Top             =   720
      Width           =   1395
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim i As Integer

Private Sub Form_Load()
i = 1
End Sub
Private Sub Timer1_Timer()
 Label1.FontSize = 12
 Label1.Caption = Mid("Developed By : Vel Murugan.S", 1, i)
 i = i + 1
 If i = Len("Developed By : Vel Murugan.S") + 1 Then i = 1
     
End Sub
