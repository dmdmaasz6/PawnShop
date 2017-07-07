VERSION 5.00
Begin VB.Form frmDBTango 
   Caption         =   "Tango Updates"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDBTango.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6855
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6375
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   550
         Left            =   2160
         TabIndex        =   3
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Select 'Update' to execute, or exit ..."
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   6135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Entries need to be updated as Tango Stock Entries"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   6135
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Exit"
      Height          =   550
      Left            =   2400
      TabIndex        =   0
      Top             =   2760
      Width           =   1935
   End
End
Attribute VB_Name = "frmDBTango"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCount As String

Private Sub cmdOK_Click()
Unload Me

End Sub

Private Sub cmdUpdate_Click()
On Error GoTo lblErr
MousePointer = vbHourglass
    
    rs.Source = "Update NormalStock set CellphoneAssessories = 2 " & _
                "where Description Like '%Tango%' and CellphoneAssessories <> 2"

    rs.Open
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
Label1.Caption = sCount & " Entries were updated as Tango Stock Entries"
Label2.Caption = "Select Exit to close this window."
cmdUpdate.Enabled = False
MousePointer = vbNormal

Exit Sub
lblErr:
MsgBox "Error Occured during update of database" & vbCrLf & "Error: " & Err.Description
cmdUpdate.Enabled = True
MousePointer = vbNormal

End Sub

Private Sub Form_Load()
MousePointer = vbHourglass
    
    rs.Source = "select count(*)" & _
                "from NormalStock where Description Like '%Tango%' and CellphoneAssessories <> 2"

    rs.Open
    
    sCount = rs.Fields(0).Value

    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    
    Label1.Caption = sCount & " Entries need to be updated as Tango Stock Entries"
    If sCount = 0 Then
        Label2.Caption = "Select Exit to close this window."
        cmdUpdate.Enabled = False
    End If
    
MousePointer = vbNormal
End Sub
