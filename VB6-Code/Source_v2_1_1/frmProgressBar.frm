VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgressBar 
   Caption         =   "Form1"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   2520
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Printing Report ...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   1860
   End
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 1 And KeyCode = 27 Then Unload Me

End Sub


Private Sub Form_Load()
Me.Caption = App.EXEName
i = 0

End Sub

Private Sub Timer1_Timer()
i = i + 1
If i > 1000 Then i = 0
ProgressBar1.Value = i

End Sub
