VERSION 5.00
Begin VB.Form frmCashRegister 
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5700
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add new Cash Register ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   840
      TabIndex        =   5
      Top             =   2520
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      Caption         =   " Used Cash Register "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   4935
      Begin VB.ComboBox cboCashRegister 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   325
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Current Cash Register "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   4935
      Begin VB.Label lblCR 
         Alignment       =   2  'Center
         Caption         =   "lblCR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   2040
      TabIndex        =   0
      Top             =   3360
      Width           =   1575
   End
End
Attribute VB_Name = "frmCashRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCashRegiter As String

Public Function SetCashRegister() As String
    Dim strCR As String
    'Get Cash Register Location
    strCR = UCase(InputBox("Type the name of the cash register to which all transactions must apply.", "Cash Register Location"))
    If Len(strCR) > 20 Then
        MsgBox "The Cash Register name may not exceed 20 characters", vbInformation, "Name too long"
        SetCashRegister = "not set"
        Exit Function
    End If
    If Len(Trim(strCR)) <> 0 Then
        'Save setting to registry
        SaveSetting "Cash&BuyBack", "Settings", "CRLocation", strCR
        'Save setting in global variable
        gstrCR = strCR
        ' Return name of selected file
        AddCashRegister strCR
        SetCashRegister = strCR
    Else 'no name specified or cancelled
        SetCashRegister = "not set"
    End If
End Function

Public Function CashRegisterExists(strCR) As Boolean
    Dim i As Integer
    Dim found As Boolean
    found = False
    strCR = UCase(strCR)
    For i = 0 To cboCashRegister.ListCount - 1
        If cboCashRegister.List(i) = strCR Then
            found = True
            'Make item selected in cbo
            cboCashRegister.ListIndex = i
        End If
    Next
    CashRegisterExists = found
End Function

Private Sub cmdAdd_Click()
    Dim strResp As String
    strResp = SetCashRegister
    If strResp = "not set" Then
        MsgBox "The cash register location has not been changed.", vbInformation, "Cash Register Unchanged"
    Else
        MsgBox "The cash register location has been successfully set to " & strResp, vbInformation, "Setting Successfull"
        gstrCR = strResp
        lblCR.Caption = gstrCR
        AddCashRegister strResp
    End If
    
End Sub

Public Sub AddCashRegister(strCR As String)
    If CashRegisterExists(strCR) <> True Then
        cboCashRegister.AddItem strCR
        'Make item selected in cbo
        cboCashRegister.ListIndex = cboCashRegister.NewIndex
    End If
End Sub

Private Sub cmdExit_Click()
    strCashRegiter = ""
    Me.Hide
    
End Sub

Private Sub cmdOK_Click()

    If cboCashRegister.Text <> "" Then
        If cboCashRegister.Text <> lblCR.Caption Then
            If MsgBox("Do you want to change the Cash Register from " & lblCR.Caption & " to " & cboCashRegister.Text & "?", vbQuestion + vbYesNo) = vbYes Then
                SaveSetting "Cash&BuyBack", "Settings", "CRLocation", cboCashRegister.Text
                MsgBox "The cash register location has been successfully set to " & cboCashRegister.Text, vbInformation, "Setting Successfull"
                gstrCR = cboCashRegister.Text
                lblCR.Caption = gstrCR
            End If
        End If
    End If
    
    strCashRegiter = lblCR.Caption
    Me.Hide
    
    
End Sub

Public Sub LoadCR()
Dim i As Integer
Dim found As Boolean

    strCashRegiter = UCase(strCashRegiter)
    For i = 0 To cboCashRegister.ListCount - 1
        If cboCashRegister.List(i) = strCashRegiter Then
            cboCashRegister.ListIndex = i
        End If
    Next
    
End Sub

