VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLayBuyPayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lay Bye Payment"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   Icon            =   "frmLayBuyPayment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   9615
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4080
      TabIndex        =   53
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   " Item Details "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   2655
      Left            =   4560
      TabIndex        =   41
      Top             =   720
      Width           =   4935
      Begin VB.TextBox txtLBID 
         DataField       =   "LBID"
         DataMember      =   "cmdLayBuyDetails"
         DataSource      =   "dePNC"
         Height          =   285
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   46
         ToolTipText     =   "Lay Bye ID"
         Top             =   240
         Width           =   1020
      End
      Begin VB.TextBox txtNSID 
         DataField       =   "NSID"
         DataMember      =   "cmdLayBuyDetails"
         DataSource      =   "dePNC"
         Height          =   285
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   45
         ToolTipText     =   "NormalStock ID"
         Top             =   615
         Width           =   1020
      End
      Begin VB.TextBox txtPSID 
         DataField       =   "PSID"
         DataMember      =   "cmdLayBuyDetails"
         DataSource      =   "dePNC"
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   44
         ToolTipText     =   "PawnStock ID"
         Top             =   600
         Width           =   1140
      End
      Begin VB.TextBox txtDescription 
         DataField       =   "Description"
         DataMember      =   "cmdLayBuyDetails"
         DataSource      =   "dePNC"
         Height          =   1005
         Left            =   1245
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   43
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtQuantity 
         DataField       =   "Quantity"
         DataMember      =   "cmdLayBuyDetails"
         DataSource      =   "dePNC"
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LB ID:"
         Height          =   195
         Index           =   5
         Left            =   645
         TabIndex        =   51
         ToolTipText     =   "Lay Bye ID"
         Top             =   285
         Width           =   450
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NS ID:"
         Height          =   195
         Index           =   6
         Left            =   615
         TabIndex        =   50
         ToolTipText     =   "NormalStock ID"
         Top             =   660
         Width           =   480
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "or    PS ID:"
         Height          =   195
         Index           =   7
         Left            =   2475
         TabIndex        =   49
         ToolTipText     =   "PawnStock ID"
         Top             =   675
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   255
         Index           =   8
         Left            =   -720
         TabIndex        =   48
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
         Height          =   255
         Index           =   13
         Left            =   -705
         TabIndex        =   47
         Top             =   1005
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Customer Details "
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
      Height          =   2655
      Left            =   120
      TabIndex        =   30
      Top             =   720
      Width           =   4335
      Begin VB.TextBox txtName 
         DataField       =   "Name"
         DataMember      =   "cmdLayBuyDetails"
         DataSource      =   "dePNC"
         Height          =   285
         Left            =   885
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   240
         Width           =   3300
      End
      Begin VB.TextBox txtIDNr 
         DataField       =   "IDNr"
         DataMember      =   "cmdLayBuyDetails"
         DataSource      =   "dePNC"
         Height          =   285
         Left            =   885
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   600
         Width           =   3300
      End
      Begin VB.TextBox txtTelH 
         DataField       =   "TelH"
         DataMember      =   "cmdLayBuyDetails"
         DataSource      =   "dePNC"
         Height          =   285
         Left            =   885
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   960
         Width           =   3300
      End
      Begin VB.TextBox txtTelW 
         DataField       =   "TelW"
         DataMember      =   "cmdLayBuyDetails"
         DataSource      =   "dePNC"
         Height          =   285
         Left            =   885
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1320
         Width           =   3300
      End
      Begin VB.TextBox txtAddress 
         DataField       =   "Address"
         DataMember      =   "cmdLayBuyDetails"
         DataSource      =   "dePNC"
         Height          =   870
         Left            =   885
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   1680
         Width           =   3300
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Index           =   0
         Left            =   -1080
         TabIndex        =   40
         Top             =   285
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IDNr:"
         Height          =   255
         Index           =   1
         Left            =   -1080
         TabIndex        =   39
         Top             =   660
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TelH:"
         Height          =   255
         Index           =   2
         Left            =   -1080
         TabIndex        =   38
         Top             =   1050
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TelW:"
         Height          =   255
         Index           =   3
         Left            =   -1080
         TabIndex        =   37
         Top             =   1425
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   255
         Index           =   4
         Left            =   -1080
         TabIndex        =   36
         Top             =   1800
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdSearchForLBID 
      Caption         =   "Search for Lay Bye ID"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame fraCancelLayBuy 
      Caption         =   "Cancel Lay Bye:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   4800
      TabIndex        =   22
      Top             =   5280
      Width           =   4695
      Begin VB.CommandButton cmdCancelLayBuy 
         Caption         =   "Cancel Lay Bye"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtStoreFee 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   2880
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblRefund 
         Alignment       =   1  'Right Justify
         Caption         =   "23"
         Height          =   255
         Left            =   2400
         TabIndex        =   28
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Refund:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblDepositAndPayments 
         Alignment       =   1  'Right Justify
         Caption         =   "23"
         Height          =   255
         Left            =   2400
         TabIndex        =   26
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Deposit and payments:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   350
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Store fee to be subtracted:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   680
         Width           =   2055
      End
   End
   Begin VB.CheckBox chkCancelLayBuy 
      Caption         =   "Cancel Lay Buy"
      Enabled         =   0   'False
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
      Left            =   7680
      TabIndex        =   21
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Frame fraNextPayment 
      Caption         =   "Next Payment:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   120
      TabIndex        =   15
      Top             =   5280
      Width           =   4575
      Begin VB.CommandButton cmdProcessPayment 
         Caption         =   "Process Payment"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpPaymentDate 
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         Format          =   58654720
         CurrentDate     =   37065
      End
      Begin VB.Label Label5 
         Caption         =   "Amount (R):"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   775
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Payment Date:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   350
         Width           =   1215
      End
   End
   Begin VB.TextBox txtAmountDue 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtTerms 
      DataField       =   "Terms"
      DataMember      =   "cmdLayBuyDetails"
      DataSource      =   "dePNC"
      Height          =   285
      Left            =   1035
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4590
      Width           =   3375
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   1  'Right Justify
      DataField       =   "Price"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataMember      =   "cmdLayBuyDetails"
      DataSource      =   "dePNC"
      Height          =   285
      Left            =   1035
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4215
      Width           =   1080
   End
   Begin VB.TextBox txtDeposit 
      Alignment       =   1  'Right Justify
      DataField       =   "Deposit"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataMember      =   "cmdLayBuyDetails"
      DataSource      =   "dePNC"
      Height          =   285
      Left            =   1035
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3825
      Width           =   1080
   End
   Begin VB.TextBox txtStartDate 
      DataField       =   "StartDate"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "d-MMM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataMember      =   "cmdLayBuyDetails"
      DataSource      =   "dePNC"
      Height          =   285
      Left            =   1035
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3450
      Width           =   2040
   End
   Begin VB.CommandButton cmdGetDetails 
      Caption         =   "Get Details"
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmLayBuyPayment.frx":08CA
      Height          =   1680
      Left            =   5400
      TabIndex        =   11
      Top             =   3480
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   2963
      _Version        =   393216
      BackColor       =   14737632
      FixedCols       =   0
      AllowUserResizing=   3
      DataMember      =   "cmdLayBuyHistory"
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   3
      _Band(0)._MapCol(0)._Name=   "LBID"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Alignment=   7
      _Band(0)._MapCol(0)._Hidden=   -1  'True
      _Band(0)._MapCol(1)._Name=   "PaymentDate"
      _Band(0)._MapCol(1)._Caption=   "Date"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "Amount"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(2)._Alignment=   7
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Due:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Payments:"
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Terms:"
      Height          =   255
      Index           =   12
      Left            =   -930
      TabIndex        =   9
      Top             =   4635
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price:"
      Height          =   195
      Index           =   11
      Left            =   480
      TabIndex        =   7
      Top             =   4260
      Width           =   405
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit:"
      Height          =   195
      Index           =   10
      Left            =   300
      TabIndex        =   5
      Top             =   3870
      Width           =   585
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "StartDate:"
      Height          =   255
      Index           =   9
      Left            =   -930
      TabIndex        =   3
      Top             =   3495
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lay Bye ID:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   1215
   End
End
Attribute VB_Name = "frmLayBuyPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private BoundControls As Collection
Private curAmountDue As Currency
Private curPaymentsDone As Currency

Private Sub chkCancelLayBuy_Click()
    'Reverse visibility of frames
    'fraCancelLayBuy.Visible = Not fraCancelLayBuy.Visible
    'fraNextPayment.Visible = Not fraNextPayment.Visible
    If fraCancelLayBuy.Enabled = True Then
        fraCancelLayBuy.Enabled = False
        cmdCancelLayBuy.Enabled = False
        fraNextPayment.Enabled = True
        'cmdProcessPayment.Enabled = True
    Else
        fraCancelLayBuy.Enabled = True
        cmdCancelLayBuy.Enabled = True
        fraNextPayment.Enabled = False
        'cmdProcessPayment.Enabled = False
    End If
End Sub

Private Sub cmdCancelLayBuy_Click()
    On Error GoTo ErrHandle
    ' Embed all updates in a transaction
    If dePNC.cnPNC.State = adStateClosed Then dePNC.cnPNC.Open
    dePNC.cnPNC.BeginTrans
    Dim CmdDE As ADODB.Command
    Set CmdDE = New ADODB.Command
    CmdDE.ActiveConnection = dePNC.cnPNC
    
    ' LayBuy table: change status to cancelled(9)
    CmdDE.CommandText = "Update LayBuy set status=9 where LBID=" & txtID.Text
    CmdDE.Execute
    ' PawnStock or NormalStock table: move quantity from QuantityLayBuy to Quantity
    If txtPSID.Text <> 0 Then 'PawnStock
        CmdDE.CommandText = "update pawnstock set QuantityLayBuy=QuantityLayBuy-" & txtQuantity.Text & " where PSID=" & txtPSID.Text
        CmdDE.Execute
    Else 'NormalStock
        CmdDE.CommandText = "update normalstock set QuantityLayBuy=QuantityLayBuy-" & txtQuantity.Text & " where NSID=" & txtNSID.Text
        CmdDE.Execute
    End If
    
    ' CashTransactions table: refund of deposit+payments-storefee
    dePNC.cmdInsertCashTransaction gstrCR, "Lay Bye Cancelled", dtpPaymentDate.Value, 0, 0, 0, 0, 0, "", -(lblRefund.Caption), 0, txtID.Text, 0

    dePNC.cnPNC.CommitTrans
    MsgBox "The client must be paid out an amount of N$" & Format(lblRefund.Caption, "#,###.00"), vbInformation, "Pay out"
    Unload Me
    Exit Sub

ErrHandle:
    MsgBox "An error occured.  This may have prevented the information you have just entered to be updated to the database.  To ensure that the database stay up to date, take the following steps:" & vbCrLf & _
        "1. On the Main Menu Window, generate a daily report." & vbCrLf & _
        "2. Try locate in the appropriate section of the report, the information you have just entered." & vbCrLf & _
        "3. If you find it, the information has been updated to the database; if not, you should try to redo the transaction." & vbCrLf & _
        "4. If you get this same error, inform the programmer.", vbCritical, "Unsuccessful database update."
    ErrorReport.WriteLine Date & " | " & Time & " | frmLayBuyPayment\cmdCancelLayBuy_Click | " & Err.Description & "(" & Err.Number & ")"
    dePNC.cnPNC.RollbackTrans
    Unload Me
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Public Sub cmdGetDetails_Click()
    ' You must close the recordset before changing the parameter.
    If dePNC.rscmdLayBuyDetails.State = adStateOpen Then
        dePNC.rscmdLayBuyDetails.Close
    End If
    ' Check that ID has been given
    If Trim(txtID.Text) = "" Then
        MsgBox "You must specify a Lay Buy ID", vbInformation, "Missing Information"
        txtID.SetFocus
        Exit Sub
    End If
    'Execute cmdLayBuyDetails with required parameters
    dePNC.cmdLayBuyDetails CLng(txtID.Text)
    'Test whether LBID is valid
    If dePNC.rscmdLayBuyDetails.EOF Then
        MsgBox "The Lay Buy ID is not valid.  It either does not exist, or has been cancelled or been completed.", vbInformation, "Invalid ID"
        txtID.Text = ""
        txtID.SetFocus
        Exit Sub
    End If
    'Refresh binding, so that content displays
    Dim ctrl As Variant
    For Each ctrl In BoundControls
        With ctrl
            .DataMember = "cmdLayBuyDetails"
            Set .DataSource = dePNC
        End With
    Next
    MSHFlexGrid1.DataMember = "cmdLayBuyHistory"
    Set MSHFlexGrid1.DataSource = dePNC
    ' Calculate and display amount still due
    AmountDue
    
    ' Make cmdProcessPayment available
    cmdProcessPayment.Enabled = True
    chkCancelLayBuy.Enabled = True
    
    ' Populate CancelLayBuy controls
    lblDepositAndPayments.Caption = curPaymentsDone + txtDeposit.Text
    lblRefund.Caption = curPaymentsDone + txtDeposit.Text
End Sub

Private Sub cmdProcessPayment_Click()
    'Check that Payment Amount has been given
    If Not CCur(txtAmount.Text) > 0 Then
        MissingInfo "a payment amount"
        txtAmount.SetFocus
        Exit Sub
    End If
    
    'Ensure that details is shown
    cmdGetDetails_Click
    
    On Error GoTo ErrHandle
    ' Embed all updates in a transaction
    If dePNC.cnPNC.State = adStateClosed Then dePNC.cnPNC.Open
    dePNC.cnPNC.BeginTrans
    
    ' Update LayBuyHistory table with latest payment
    Dim CmdDE As ADODB.Command
    Set CmdDE = New ADODB.Command
    With CmdDE
        .ActiveConnection = dePNC.cnPNC
        .CommandType = adCmdText
        .CommandText = "Insert into laybuyhistory(LBID, PaymentDate, Amount) values(" & txtLBID.Text & ", #" & dtpPaymentDate.Value & "#," & txtAmount.Text & ")"
        .Execute
    End With
    
    ' Get this last LBHID
    Dim RsDE As ADODB.Recordset
    Set RsDE = New ADODB.Recordset
    RsDE.Open "Select max(LBHID) as maxLBHID from LayBuyHistory", dePNC.cnPNC
    ' Insert into CashTransactions
    dePNC.cmdInsertCashTransaction gstrCR, "Lay Bye Payment", dtpPaymentDate.Value, 0, 0, 0, 0, 0, "", txtAmount.Text, 0, 0, RsDE.Fields("maxLBHID").Value
    RsDE.Close
    
    ' Print Report
    ' You must close the recordset before changing the parameter.
    If dePNC.rscmdLayBuyDetails.State = adStateOpen Then
        dePNC.rscmdLayBuyDetails.Close
    End If
    'Execute cmdLayBuyDetails with required parameters
    dePNC.cmdLayBuyDetails CLng(txtID.Text)
    With drLayBuyFullReport
        .DataMember = "cmdLayBuyDetails"
        Set .DataSource = dePNC
        .Refresh
    End With
    drLayBuyFullReport.Show vbModal
    'drlaybuyfullreport.PrintReport
    
    ' Check if this was final payment
    If CDec(txtAmountDue.Text) <= CDec(txtAmount.Text) Then
        MsgBox "This was the final payment.  Client may take stock item.", vbInformation, "Lay Bye paid in full"
        'Move quantity from QuantityLayBuy to QuantitySold
        If txtPSID.Text <> 0 Then 'PawnStock
            CmdDE.CommandText = "update pawnstock set QuantityLayBuy=QuantityLayBuy-" & txtQuantity.Text & ", QuantitySold=QuantitySold+" & txtQuantity.Text & " where PSID=" & txtPSID.Text
            CmdDE.Execute
        Else 'NormalStock
            CmdDE.CommandText = "update normalstock set QuantityLayBuy=QuantityLayBuy-" & txtQuantity.Text & ", QuantitySold=QuantitySold+" & txtQuantity.Text & " where NSID=" & txtNSID.Text
            CmdDE.Execute
        End If
        'Change status of LayBuy to completed (8)
        CmdDE.CommandText = "update laybuy set status=8 where LBID=" & txtLBID.Text
        CmdDE.Execute
    End If
    dePNC.cnPNC.CommitTrans
    Unload Me
    Exit Sub

ErrHandle:
    MsgBox "An error occured.  This may have prevented the information you have just entered to be updated to the database.  To ensure that the database stay up to date, take the following steps:" & vbCrLf & _
        "1. On the Main Menu Window, generate a daily report." & vbCrLf & _
        "2. Try locate in the appropriate section of the report, the information you have just entered." & vbCrLf & _
        "3. If you find it, the information has been updated to the database; if not, you should try to redo the transaction." & vbCrLf & _
        "4. If you get this same error, inform the programmer.", vbCritical, "Unsuccessful database update."
    ErrorReport.WriteLine Date & " | " & Time & " | frmLayBuyPayment\cmdProcessPayment_Click | " & Err.Description & "(" & Err.Number & ")"
    dePNC.cnPNC.RollbackTrans
    Unload Me
End Sub

Private Sub cmdSearchForLBID_Click()
    frmLayBuySearch.Show vbModal
End Sub

Private Sub Form_Load()
    Set BoundControls = New Collection
    BoundControls.Add txtName
    BoundControls.Add txtIDNr
    BoundControls.Add txtTelH
    BoundControls.Add txtTelW
    BoundControls.Add txtAddress
    BoundControls.Add txtStartDate
    BoundControls.Add txtDeposit
    BoundControls.Add txtTerms
    BoundControls.Add txtPrice
    BoundControls.Add txtLBID
    BoundControls.Add txtPSID
    BoundControls.Add txtNSID
    BoundControls.Add txtDescription
    BoundControls.Add txtQuantity
    dtpPaymentDate.Value = Date
End Sub

Private Sub AmountDue()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "Select sum(amount) as Total from LayBuyHistory where LBID=" & txtLBID.Text, dePNC.cnPNC
    If IsNull(rs.Fields("total").Value) Then
        curPaymentsDone = 0
    Else
        curPaymentsDone = rs.Fields("total").Value
    End If
    curAmountDue = dePNC.rscmdLayBuyDetails.Fields("Price").Value - dePNC.rscmdLayBuyDetails.Fields("Deposit").Value - curPaymentsDone
    rs.Close
    'txtAmount.Text = curAmountDue
    txtAmountDue.Text = Format(curAmountDue, "#,##0.00")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Close any open recordsets
    If dePNC.rscmdLayBuyDetails.State = adStateOpen Then
        dePNC.rscmdLayBuyDetails.Close
    End If
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    'Ensure that only numbers, backspace or decimal point are entered
    If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 46) Then
        MsgBox "Amount must have a numeric value.", vbInformation, "Key invalid"
        KeyAscii = 0
    End If
End Sub

Private Sub txtAmount_LostFocus()
    SetCurr txtAmount
End Sub

Private Sub txtStoreFee_Change()
    If IsNumeric(txtStoreFee.Text) Then
        lblRefund.Caption = lblDepositAndPayments.Caption - txtStoreFee.Text
    End If
End Sub
