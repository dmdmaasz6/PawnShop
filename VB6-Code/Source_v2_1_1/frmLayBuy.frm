VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLayBuy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Lay Bye Transaction "
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   Icon            =   "frmLayBuy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   9330
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLayBuy 
      Caption         =   "Lay Bye"
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   7200
      TabIndex        =   33
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame fraItems 
      Caption         =   " LayBye Item: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2520
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cboPawnStockID 
         Height          =   315
         Index           =   0
         ItemData        =   "frmLayBuy.frx":08CA
         Left            =   1200
         List            =   "frmLayBuy.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "PawnStock ID"
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboNormalStockID 
         Height          =   315
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "NormalStock ID"
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cboQuantity 
         Height          =   315
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtDescription 
         Height          =   315
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtSerialNr 
         Height          =   315
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtPrice 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   3
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtSubTotal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   1035
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   31
         Top             =   1410
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PS ID:"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   30
         ToolTipText     =   "PawnStock ID"
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NS ID:"
         Height          =   255
         Left            =   -120
         TabIndex        =   29
         ToolTipText     =   "NormalStock ID"
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SerialNr:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1755
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Price:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   2115
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Price x Quantity:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   26
         Top             =   2100
         Width           =   1575
      End
   End
   Begin VB.Frame fraLayBuyTransaction 
      Caption         =   " Lay Bye Transaction: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1695
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   6135
      Begin VB.TextBox txtTerms 
         Height          =   765
         Left            =   3720
         MaxLength       =   100
         TabIndex        =   11
         Top             =   780
         Width           =   2055
      End
      Begin VB.TextBox txtDeposit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   780
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20578304
         CurrentDate     =   37065
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Terms:"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   21
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit:"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   850
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date:"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   400
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Client:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   2535
      Left            =   5280
      TabIndex        =   12
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtIDnr 
         Height          =   315
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtTelW 
         Height          =   315
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtTelH 
         Height          =   315
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   7
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtAddress 
         Height          =   765
         Left            =   1320
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ID number:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tel(w):"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   1035
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tel(h):"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1410
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Full Address:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmLayBuy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
    dtpDate.Value = Date
    CreateIDs (0)
End Sub

Private Sub cboNormalStockID_Click(Index As Integer)
    'Ignore if text is zero
    If cboNormalStockID(Index).Text = "0" Then
        If cboPawnStockID(Index).Text = "0" Then
            'Both IDs empty.  Clear controls.
            ClearControls Index
            Exit Sub
        Else
            'Only NormalStockID is 0
            Exit Sub
        End If
    End If
    
    ' Neutralize cboNormalStockID, if it had been selected
    cboPawnStockID(Index).ListIndex = 0
    
    'Retrieve relevant information and populate controls
    rs.Open "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, RecommendedSellingPrice From NormalStock WHERE NSID=" & cboNormalStockID(Index).Text
    'Empty quantity combobox and then repopulate
    cboQuantity(Index).Clear
    Dim i As Integer
    'cboQuantity(Index).AddItem "0"
    For i = 1 To rs.Fields("Quantity").Value
        cboQuantity(Index).AddItem i
    Next i
    cboQuantity(Index).ListIndex = 0
    'Populate other controls
    txtDescription(Index).Text = rs.Fields("Description").Value
    txtPrice(Index).Text = rs.Fields("RecommendedSellingPrice").Value
    rs.Close
End Sub

Private Sub cboPawnStockID_Click(Index As Integer)
    'Ignore if text is zero
    If cboPawnStockID(Index).Text = "0" Then
        If cboNormalStockID(Index).Text = "0" Then
            'Both IDs empty.  Clear controls.
            ClearControls Index
            Exit Sub
        Else
            'Only PawnStockID is 0
            Exit Sub
        End If
    End If
    
    ' Neutralize cboNormalStockID, if it had been selected
    cboNormalStockID(Index).ListIndex = 0
    
    'Retrieve relevant information and populate controls
    rs.Open "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, SerialNr, Price From PawnStock WHERE PawnStock.PSID=" & cboPawnStockID(Index).Text
    'Empty quantity combobox and then repopulate
    cboQuantity(Index).Clear
    Dim i As Integer
    'cboQuantity(Index).AddItem "0"
    For i = 1 To rs.Fields("Quantity").Value
        cboQuantity(Index).AddItem i
    Next i
    cboQuantity(Index).ListIndex = 0
    'Populate other controls
    txtDescription(Index).Text = rs.Fields("Description").Value
    txtSerialNr(Index).Text = rs.Fields("SerialNr").Value
    txtPrice(Index).Text = rs.Fields("Price").Value
    rs.Close
End Sub

Private Sub CreateIDs(Index As Integer)
    'Populate comboboxes
    '[PawnStock]
    cboPawnStockID(Index).AddItem "0"
    rs.Open "Select PSID from PawnStock where (Status=3) and (Quantity-QuantitySold-QuantityLayBuy>0) order by PSID"
    Do Until rs.EOF
        cboPawnStockID(Index).AddItem rs.Fields("PSID").Value
        rs.MoveNext
    Loop
    rs.Close
    'Make first item (namely zero) selected
    cboPawnStockID(Index).ListIndex = 0
    '[NormalStock]
    cboNormalStockID(Index).AddItem "0"
    rs.Open "Select NSID from NormalStock where (Quantity-QuantitySold-QuantityLayBuy>0) order by NSID"
    Do Until rs.EOF
        cboNormalStockID(Index).AddItem rs.Fields("NSID").Value
        rs.MoveNext
    Loop
    rs.Close
    'Make first item (namely zero) selected
    cboNormalStockID(Index).ListIndex = 0
End Sub

Private Sub cboQuantity_Click(Index As Integer)
    CalculateSubTotal (Index)
End Sub

Private Sub CalculateSubTotal(Index As Integer)
    'Skip if values missing
    If cboQuantity(Index).Text = "" Or txtPrice(Index).Text = "" Then
        Exit Sub
    End If
    txtSubTotal(Index).Text = cboQuantity(Index).Text * txtPrice(Index).Text
End Sub

Private Sub txtDeposit_KeyPress(KeyAscii As Integer)
    'Ensure that only numbers, backspace or decimal point are entered
    If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 46) Then
        MsgBox "Deposit must have a numeric value.", vbInformation, "Key invalid"
        KeyAscii = 0
    End If
End Sub

Private Sub txtDeposit_LostFocus()
    SetCurr txtDeposit
End Sub

Private Sub txtPrice_Change(Index As Integer)
    CalculateSubTotal (Index)
End Sub

Private Sub ClearControls(Index As Integer)
    'Clear Controls
    cboQuantity(Index).Clear
    txtDescription(Index).Text = ""
    txtSerialNr(Index).Text = ""
    txtPrice(Index).Text = ""
    txtSubTotal(Index).Text = ""
    'CalculateSubTotal (Index)
End Sub

Private Sub cmdLayBuy_Click()
    ' Ensure that minimun information has been specified
    If Trim(txtName.Text) = "" Then
        MissingInfo "the client name"
        txtName.SetFocus
        Exit Sub
    ElseIf Trim(txtDeposit.Text) = "" Then
        MissingInfo "the deposit amount"
        txtDeposit.SetFocus
        Exit Sub
    ElseIf Trim(txtTerms.Text) = "" Then
        'MissingInfo "the terms and conditions"
        'txtTerms.SetFocus
        'Exit Sub
    End If
    
    ' Embed updates to PawnStock, NormalStock and LayBuy tables in a transaction
    On Error GoTo ErrorTrap
    If dePNC.cnPNC.State = adStateClosed Then dePNC.cnPNC.Open
    dePNC.cnPNC.BeginTrans
    
    ' Update PawnStock and NormalStock QuantityLayBuy fields
    Dim CmdDE As ADODB.Command
    Set CmdDE = New ADODB.Command
    Dim RsDE As ADODB.Recordset
    Set RsDE = New ADODB.Recordset
    CmdDE.ActiveConnection = dePNC.cnPNC
    CmdDE.CommandType = adCmdText
    If cboPawnStockID(i).Text <> "0" Then
        '//Process PawnStock LayBuy//
        '**************************
        'Update to PawnStock Table
        CmdDE.CommandText = "Update PawnStock set QuantityLayBuy=QuantityLayBuy+" & cboQuantity(i).Text & " where PSID=" & cboPawnStockID(i).Text
        CmdDE.Execute
        'Insert into LayBuy
        dePNC.cmdNewLayBuy txtName.Text, txtTelW.Text, txtTelH.Text, txtAddress.Text, txtIDNr.Text, cboPawnStockID(0).Text, 0, dtpDate.Value, txtDeposit.Text, txtTerms.Text, txtSubTotal(0).Text, txtDescription(0).Text, cboQuantity(0).Text
        'Determine ID of current LayBuy
        Dim ThisLBID As Long
        RsDE.Open "Select max(LBID) as MaxLBID from LayBuy", dePNC.cnPNC
        ThisLBID = RsDE.Fields("MaxLBID").Value
        RsDE.Close
        'Update to CashTransactions (deposit)
        dePNC.cmdInsertCashTransaction gstrCR, "Lay Bye Deposit", dtpDate.Value, 0, cboPawnStockID(i).Text, cboQuantity(i).Text, 0, 0, txtDescription(i).Text, txtSubTotal(i).Text, txtPrice(i).Text, ThisLBID, 0
    ElseIf cboNormalStockID(i).Text <> "0" Then
        '//Process NormalStock LayBuy//
        '**************************
        'Update to NormalStock Table
        CmdDE.CommandText = "Update NormalStock set QuantityLayBuy=QuantityLayBuy+" & cboQuantity(i).Text & " where NSID=" & cboNormalStockID(i).Text
        CmdDE.Execute
        'Insert into LayBuy
        dePNC.cmdNewLayBuy txtName.Text, txtTelW.Text, txtTelH.Text, txtAddress.Text, txtIDNr.Text, 0, cboNormalStockID(0).Text, dtpDate.Value, txtDeposit.Text, txtTerms.Text, txtSubTotal(0).Text, txtDescription(0).Text, cboQuantity(0).Text
        'Determine ID of current LayBuy
        RsDE.Open "Select max(LBID) as MaxLBID from LayBuy", dePNC.cnPNC
        ThisLBID = RsDE.Fields("MaxLBID").Value
        RsDE.Close
        'Update to CashTransactions (deposit)
        dePNC.cmdInsertCashTransaction gstrCR, "Lay Bye Deposit", dtpDate.Value, 0, 0, cboQuantity(i).Text, cboNormalStockID(i).Text, 0, txtDescription(i).Text, txtSubTotal(i).Text, txtPrice(i).Text, ThisLBID, 0
    Else
        'Neither PawnStockID or NormalStockID specified
        MissingInfo "an item to lay buy"
        cboPawnStockID(i).SetFocus
        Exit Sub
    End If
    dePNC.cnPNC.CommitTrans
    
    ' Print Invoice/Report
    drNewLayBuyReport.Show vbModal
    'drnewlaybuyreport.PrintReport
    
    Unload Me
    Exit Sub
    
ErrorTrap:
    MsgBox "An error occured.  This may have prevented the information you have just entered to be updated to the database.  To ensure that the database stay up to date, take the following steps:" & vbCrLf & _
        "1. On the Main Menu Window, generate a daily report." & vbCrLf & _
        "2. Try locate in the appropriate section of the report, the information you have just entered." & vbCrLf & _
        "3. If you find it, the information has been updated to the database; if not, you should try to redo the transaction." & vbCrLf & _
        "4. If you get this same error, inform the programmer.", vbCritical, "Unsuccessful database update."
    ErrorReport.WriteLine Date & " | " & Time & " | frmLayBuy\cmdLayBuy_Click | " & Err.Description & "(" & Err.Number & ")"
    dePNC.cnPNC.RollbackTrans
    Unload Me
End Sub

Private Sub txtPrice_KeyPress(Index As Integer, KeyAscii As Integer)
    'Ensure that only numbers, backspace or decimal point are entered
    If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 46) Then
        MsgBox "Price must have a numeric value.", vbInformation, "Key invalid"
        KeyAscii = 0
    End If

End Sub

Private Sub txtPrice_LostFocus(Index As Integer)
    SetCurr txtPrice(Index)
End Sub

Private Sub txtSubTotal_Change(Index As Integer)
    SetCurr txtSubTotal(Index)
End Sub

