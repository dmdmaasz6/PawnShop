VERSION 5.00
Begin VB.Form frmAddress 
   Caption         =   "Address"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddress.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   5955
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtAddress3 
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   4335
   End
   Begin VB.TextBox txtAddress2 
      Height          =   360
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin VB.TextBox txtAddress1 
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Address Line1"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   1125
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Address Line1"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   655
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Address Line1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdOK_Click()
    Dim strResp As String
    
    If Trim(txtAddress1.Text) = "" Then
        txtAddress1.Text = "not set"
        txtAddress2.Text = ""
        txtAddress3.Text = ""
    End If
    gstrAddress1 = SetAddress1(Trim(txtAddress1.Text))
    gstrAddress2 = SetAddress2(Trim(txtAddress2.Text))
    gstrAddress3 = SetAddress3(Trim(txtAddress3.Text))

    
    MsgBox "The address has been successfully changed", vbInformation, "Setting Successfull"
    gstrAddress = gstrAddress1 & " " & gstrAddress2 & " " & gstrAddress3
    Unload Me
    
End Sub

Private Sub Form_Load()
txtAddress1.Text = gstrAddress1
txtAddress2.Text = gstrAddress2
txtAddress3.Text = gstrAddress3

End Sub

Public Function SetAddress1(sIn As String) As String
        SaveSetting "Cash&BuyBack", "Settings", "Address1", sIn
        SetAddress1 = sIn
End Function

Public Function SetAddress2(sIn As String) As String
        SaveSetting "Cash&BuyBack", "Settings", "Address2", sIn
        SetAddress2 = sIn
End Function

Public Function SetAddress3(sIn As String) As String
        SaveSetting "Cash&BuyBack", "Settings", "Address3", sIn
        SetAddress3 = sIn
End Function

