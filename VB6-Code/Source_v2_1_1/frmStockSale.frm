VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStockSale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Sale"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStockSale.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   5760
      TabIndex        =   24
      Top             =   5280
      Width           =   1550
   End
   Begin VB.CommandButton cmdStockSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   120
      Width           =   1550
   End
   Begin VB.CommandButton cmdTransact 
      Caption         =   "Transact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5280
      Width           =   1550
   End
   Begin VB.Frame fraItems 
      Caption         =   " Items: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1800
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10335
      Begin VB.CheckBox chkCellphoneAssessories 
         Caption         =   "Check1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   8160
         TabIndex        =   21
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtGrandTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1225
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddAnotherItem 
         Caption         =   "Add Another Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtSubTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   700
         Width           =   1575
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   7200
         TabIndex        =   12
         Top             =   700
         Width           =   855
      End
      Begin VB.TextBox txtSerialNr 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   700
         Width           =   735
      End
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   700
         Width           =   2535
      End
      Begin VB.ComboBox cboQuantity 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   700
         Width           =   735
      End
      Begin VB.ComboBox cboNormalStockID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "NormalStock ID"
         Top             =   700
         Width           =   855
      End
      Begin VB.ComboBox cboPawnStockID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         ItemData        =   "frmStockSale.frx":08CA
         Left            =   120
         List            =   "frmStockSale.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "PawnStock ID"
         Top             =   700
         Width           =   975
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "SS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   22
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblGrandTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   17
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "SubTotal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8520
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "SerialNr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NS ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         ToolTipText     =   "NormalStock ID"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PS ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "PawnStock ID"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   360
      Left            =   2520
      TabIndex        =   18
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20709376
      CurrentDate     =   37065
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   19
      Top             =   300
      Width           =   855
   End
End
Attribute VB_Name = "frmStockSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public row As Integer 'also used from frmStockSearch
Private TopVal As Integer

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
    rs.Open "SELECT (Quantity-QuantitySold-QuantityLayBuy) as Quantity, Description, RecommendedSellingPrice, CellphoneAssessories From NormalStock WHERE NSID=" & cboNormalStockID(Index).Text
    'Empty quantity combobox and then repopulate
    cboQuantity(Index).Clear
    Dim i As Integer
    For i = 1 To rs.Fields("Quantity").Value
        cboQuantity(Index).AddItem i
    Next i
    cboQuantity(Index).ListIndex = 0
    'Populate other controls
    txtDescription(Index).Text = rs.Fields("Description").Value
    txtPrice(Index).Text = rs.Fields("RecommendedSellingPrice").Value
    SetCurr txtPrice(Index)
    txtPrice(Index).Enabled = True
    chkCellphoneAssessories(Index).Value = rs.Fields("CellphoneAssessories").Value
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
    For i = 1 To rs.Fields("Quantity").Value
        cboQuantity(Index).AddItem i
    Next i
    cboQuantity(Index).ListIndex = 0
    'Populate other controls
    txtDescription(Index).Text = rs.Fields("Description").Value
    txtSerialNr(Index).Text = rs.Fields("SerialNr").Value
    txtPrice(Index).Text = rs.Fields("Price").Value
    SetCurr txtPrice(Index)
    txtPrice(Index).Enabled = True
    rs.Close
End Sub

Private Sub cboQuantity_Click(Index As Integer)
    CalculateSubTotal (Index)
End Sub

Public Sub cmdAddAnotherItem_Click()
    row = row + 1
    'Initialize for first value
    If TopVal = 0 Then TopVal = 700
    TopVal = TopVal + 400
    CreateControl cboPawnStockID(row), TopVal
    CreateControl cboNormalStockID(row), TopVal
    CreateControl cboQuantity(row), TopVal
    CreateControl txtDescription(row), TopVal
    CreateControl txtSerialNr(row), TopVal
    CreateControl txtPrice(row), TopVal
    txtPrice(row).Enabled = False
    CreateControl txtSubTotal(row), TopVal
    CreateControl chkCellphoneAssessories(row), TopVal
    'Lower commandbutton, frame, label and textbox
    cmdAddAnotherItem.Top = TopVal + 400
    lblGrandTotal.Top = TopVal + 400
    txtGrandTotal.Top = TopVal + 400
    fraItems.Height = TopVal + 900
    'cmdTransact.top = fraItems.top + fraItems.Height + 200
    'Populate comboboxes
    CreateIDs (row)
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdStockSearch_Click()
    'frmStockSearch.fraStockSell.Visible = True
    frmStockSearch.Show vbModal
End Sub

Private Sub cmdTransact_Click()
'Update NormalStock set CellphoneAssessories = 2
    Dim sLastCmd As String
    Dim blnAtLeastOne As Boolean 'Indicates whether at least one item is being bought, display report
    blnAtLeastOne = False
    Dim i As Integer
    'Embed all updates in a transaction
    On Error GoTo ErrHandle
    sLastCmd = ""
    cn.BeginTrans
    For i = 0 To row
        If cboPawnStockID(i).Text <> "0" Then
            '//Process PawnStock Sell//
            '**************************
            'Update to PawnStock
            cmd.CommandText = "Update PawnStock set QuantitySold=QuantitySold+" & cboQuantity(i).Text & " where PSID=" & cboPawnStockID(i).Text
            sLastCmd = cmd.CommandText
            cmd.Execute
            'Update to CashTransactions
            dePNC.cmdInsertCashTransaction gstrCR, "Pawn Stock Sale", dtpDate.Value, 0, cboPawnStockID(i).Text, cboQuantity(i).Text, 0, 0, txtDescription(i).Text, txtSubTotal(i).Text, txtPrice(i).Text, 0, 0
            sLastCmd = "cmdInsertCashTransaction" & "," & gstrCR & "," & "Pawn Stock Sale" & "," & dtpDate.Value & "," & 0 & "," & cboPawnStockID(i).Text & "," & cboQuantity(i).Text & "," & 0 & "," & 0 & "," & txtDescription(i).Text & "," & txtSubTotal(i).Text & "," & txtPrice(i).Text & "," & 0 & "," & 0
            'Insert into StockSale
            dePNC.cmdInsertIntoStockSale dtpDate.Value, "PS_ID: " & cboPawnStockID(i).Text, cboQuantity(i).Text, txtDescription(i).Text, txtSerialNr(i).Text, txtPrice(i).Text, txtSubTotal(i).Text
            sLastCmd = "cmdInsertIntoStockSale" & "," & dtpDate.Value & "," & "PS_ID: " & cboPawnStockID(i).Text & "," & cboQuantity(i).Text & "," & txtDescription(i).Text & "," & txtSerialNr(i).Text & "," & txtPrice(i).Text & "," & txtSubTotal(i).Text
            blnAtLeastOne = True
            SetCBBIDStatus cboPawnStockID(i).Text
        ElseIf cboNormalStockID(i).Text <> "0" Then
            '//Process NormalStock Sell//
            '***************************
            'Update to NormalStock
            cmd.CommandText = "Update NormalStock set QuantitySold=QuantitySold+" & cboQuantity(i).Text & " where NSID=" & cboNormalStockID(i).Text
            sLastCmd = cmd.CommandText
            cmd.Execute
            'Update to CashTransactions
            If chkCellphoneAssessories(i).Value Then
                'Standard Stock
                If IsTangoStock(txtDescription(i).Text) = True Then
                    dePNC.cmdInsertCashTransaction gstrCR, "Tango Stock Sale", dtpDate.Value, 0, 0, cboQuantity(i).Text, cboNormalStockID(i).Text, 0, txtDescription(i).Text, txtSubTotal(i).Text, txtPrice(i).Text, 0, 0
                    sLastCmd = "cmdInsertCashTransaction1" & "," & "Tango Stock Sale" & "," & dtpDate.Value & "," & 0 & "," & 0 & "," & cboQuantity(i).Text & "," & cboNormalStockID(i).Text & "," & 0 & "," & txtDescription(i).Text & "," & txtSubTotal(i).Text & "," & txtPrice(i).Text & "," & 0 & "," & 0
                Else
                    dePNC.cmdInsertCashTransaction gstrCR, "Standard Stock Sale", dtpDate.Value, 0, 0, cboQuantity(i).Text, cboNormalStockID(i).Text, 0, txtDescription(i).Text, txtSubTotal(i).Text, txtPrice(i).Text, 0, 0
                    sLastCmd = "cmdInsertCashTransaction2" & "," & gstrCR & "," & "Standard Stock Sale" & "," & dtpDate.Value & "," & 0 & "," & 0 & "," & cboQuantity(i).Text & "," & cboNormalStockID(i).Text & "," & 0 & "," & txtDescription(i).Text & "," & txtSubTotal(i).Text & "," & txtPrice(i).Text & "," & 0 & "," & 0
                End If
            Else
                'Normal Stock
                If IsTangoStock(txtDescription(i).Text) = True Then
                    dePNC.cmdInsertCashTransaction gstrCR, "Tango Stock Sale", dtpDate.Value, 0, 0, cboQuantity(i).Text, cboNormalStockID(i).Text, 0, txtDescription(i).Text, txtSubTotal(i).Text, txtPrice(i).Text, 0, 0
                    sLastCmd = "cmdInsertCashTransaction3" & "," & gstrCR & "," & "Tango Stock Sale" & "," & dtpDate.Value & "," & 0 & "," & 0 & "," & cboQuantity(i).Text & "," & cboNormalStockID(i).Text & "," & 0 & "," & txtDescription(i).Text & "," & txtSubTotal(i).Text & "," & txtPrice(i).Text & "," & 0 & "," & 0
                Else
                    dePNC.cmdInsertCashTransaction gstrCR, "Normal Stock Sale", dtpDate.Value, 0, 0, cboQuantity(i).Text, cboNormalStockID(i).Text, 0, txtDescription(i).Text, txtSubTotal(i).Text, txtPrice(i).Text, 0, 0
                    sLastCmd = "cmdInsertCashTransaction4" & "," & "Normal Stock Sale" & "," & dtpDate.Value & "," & 0 & "," & 0 & "," & cboQuantity(i).Text & "," & cboNormalStockID(i).Text & "," & 0 & "," & txtDescription(i).Text & "," & txtSubTotal(i).Text & "," & txtPrice(i).Text & "," & 0 & "," & 0
                End If
            End If
            'Insert into StockSale
            dePNC.cmdInsertIntoStockSale dtpDate.Value, "NS_ID: " & cboNormalStockID(i).Text, cboQuantity(i).Text, txtDescription(i).Text, txtSerialNr(i).Text, txtPrice(i).Text, txtSubTotal(i).Text
            sLastCmd = "cmdInsertIntoStockSale" & "," & dtpDate.Value & "," & "NS_ID: " & cboNormalStockID(i).Text & "," & cboQuantity(i).Text & "," & txtDescription(i).Text & "," & txtSerialNr(i).Text & "," & txtPrice(i).Text & "," & txtSubTotal(i).Text
            blnAtLeastOne = True
        End If
    Next
    
    cn.CommitTrans
    sLastCmd = "CommitTrans"

    ' Print Invoice/Report
    If blnAtLeastOne Then drStockSale.Show vbModal
    'drstocksale.PrintReport

    'Clear StockSale table
    cmd.CommandText = "Delete from StockSale"
    sLastCmd = "Delete from StockSale"
    cmd.Execute
    
    Unload frmStockSale
    Exit Sub

ErrHandle:
    MsgBox "An error occured.  This may have prevented the information you have just entered to be updated to the database.  To ensure that the database stay up to date, take the following steps:" & vbCrLf & _
        "1. On the Main Menu Window, generate a daily report." & vbCrLf & _
        "2. Try locate in the appropriate section of the report, the information you have just entered." & vbCrLf & _
        "3. If you find it, the information has been updated to the database; if not, you should try to redo the transaction." & vbCrLf & _
        "4. If you get this same error, inform the programmer.", vbCritical, "Unsuccessful database update."
    ErrorReport.WriteLine Date & " | " & Time & " | frmStockSale\cmdTransact_Click | " & Err.Description & "(" & Err.Number & ")"
    ErrorReport.WriteLine Date & " | " & Time & " | frmStockSale\cmdTransact_Click | LastCommand(" & sLastCmd & ")"
    cn.RollbackTrans
    Unload Me

End Sub

Private Function IsTangoStock(sStock_Description As String) As Boolean

    If InStr(1, UCase(sStock_Description), "TANGO", vbTextCompare) Then
        IsTangoStock = True
    Else
        IsTangoStock = False
    End If
    
End Function

Private Sub Form_Load()
    dtpDate.Value = Date
    CreateIDs (0)
End Sub

Private Sub CreateIDs(Index As Integer)
    'Populate comboboxes
    cboPawnStockID(Index).AddItem "0"
    rs.Open "Select PSID from PawnStock where (Status=3) and (Quantity-QuantitySold-QuantityLayBuy>0) order by PSID"
    Do Until rs.EOF
        cboPawnStockID(Index).AddItem rs.Fields("PSID").Value
        rs.MoveNext
    Loop
    rs.Close
    'Make first item (namely zero) selected
    cboPawnStockID(Index).ListIndex = 0
    
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

Private Sub CalculateSubTotal(Index As Integer)
    'Skip if values missing
    If cboQuantity(Index).Text = "" Or txtPrice(Index).Text = "" Then
        Exit Sub
    End If
    txtSubTotal(Index).Text = cboQuantity(Index).Text * txtPrice(Index).Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Reset counters back to zero
    row = 0
    TopVal = 0
    'Reset recordset for drStockSale
    If dePNC.rscmdStockSaleReport.State = adStateOpen Then
        dePNC.rscmdStockSaleReport.Close
    End If
End Sub

Private Sub txtGrandTotal_Change()
    SetCurr txtGrandTotal
End Sub

Private Sub txtGrandTotal_LostFocus()
    SetCurr txtGrandTotal
End Sub

Private Sub txtPrice_Change(Index As Integer)
    'SetCurr txtPrice(Index)
    CalculateSubTotal (Index)
End Sub

Private Sub CreateControl(ctrl As Control, TopValue As Integer)
    Load ctrl
    ctrl.Top = TopValue
    If TypeOf ctrl Is TextBox Then ctrl.Text = ""
    ctrl.Visible = True
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
    Dim i As Integer
    Dim dblTotal As Currency
    For i = 0 To row
        If txtSubTotal(i).Text <> "" Then
            dblTotal = dblTotal + txtSubTotal(i).Text
        End If
    Next
    txtGrandTotal = dblTotal
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

Private Sub txtSubTotal_LostFocus(Index As Integer)
    SetCurr txtSubTotal(Index)
End Sub

Private Function SetCBBIDStatus(sPSID As String)
Dim sPnCID As String
Dim arrPnCIDStock(100) As String
Dim iBuy As Integer
Dim iSold As Integer
Dim iPnc As Integer
Dim iSQL As Integer

iPnc = 0

    rs.Source = "Select PnCID from PawnStock where PSID=" & sPSID
    rs.Open
    
    If Not rs.EOF Then
        sPnCID = rs.Fields("PnCID").Value
    End If
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    rs.Source = "Select PSID from PawnStock where PnCID=" & sPnCID
    rs.Open
    Do Until rs.EOF
        arrPnCIDStock(iPnc) = rs.Fields("PSID").Value
        iPnc = iPnc + 1
        rs.MoveNext
    Loop
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
        
    For iSQL = 0 To iPnc - 1
        rs.Source = "Select Quantity , QuantitySold from PawnStock where PSID=" & arrPnCIDStock(iSQL)
        rs.Open
        
        If Not rs.EOF Then
            iBuy = rs.Fields("Quantity").Value
            iSold = rs.Fields("QuantitySold").Value
            If iBuy > iSold Then GoTo lblExit
        End If
        
        If rs.State = adStateOpen Then
            rs.Close
        End If
    Next iSQL

    rs.Source = "update PawningTransactions set Status = 2 where PnCID = " & sPnCID
    rs.Open
    
    If rs.State = adStateOpen Then
        rs.Close
    End If

lblExit:
    If rs.State = adStateOpen Then
        rs.Close
    End If

End Function

