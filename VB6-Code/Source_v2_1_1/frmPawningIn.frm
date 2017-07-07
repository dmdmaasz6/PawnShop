VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPawningIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pawning In"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12495
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPawningIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12495
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   550
      Left            =   9840
      TabIndex        =   16
      Top             =   8085
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   " Goods "
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
      ForeColor       =   &H00C00000&
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   6000
      Width           =   3495
      Begin VB.Label lblGoods 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Pawning Transaction "
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
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   3495
      Begin VB.Label lblPawningTransaction 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Customer Information "
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
      ForeColor       =   &H00C00000&
      Height          =   2655
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   3495
      Begin VB.Label lblCustomerInformation 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdSearchForPNCID 
      Caption         =   "Search for PNCID"
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   " Search for CCB ID: "
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
      Height          =   7575
      Left            =   3840
      TabIndex        =   8
      ToolTipText     =   "Cash & BuyBack Customer ID"
      Top             =   120
      Width           =   8535
      Begin VB.Frame Frame5 
         Caption         =   " Customer Information "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   8175
         Begin VB.CommandButton cmdSearchForPNC 
            Caption         =   "Search"
            Height          =   550
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtName 
            Height          =   360
            Left            =   960
            TabIndex        =   25
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtIDNr 
            Height          =   360
            Left            =   960
            TabIndex        =   24
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtCustID 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4920
            TabIndex        =   23
            Top             =   960
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "ID:"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   675
            Width           =   735
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "CustomerID:"
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
            Left            =   3840
            TabIndex        =   27
            Top             =   1005
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgPNCs 
         Bindings        =   "frmPawningIn.frx":08CA
         Height          =   2775
         Left            =   120
         TabIndex        =   17
         Top             =   4680
         Visible         =   0   'False
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4895
         _Version        =   393216
         BackColor       =   14737632
         Cols            =   9
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         DataMember      =   "cmdSearchPNC2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   2
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0)._NumMapCols=   5
         _Band(0)._MapCol(0)._Name=   "PnCID"
         _Band(0)._MapCol(0)._RSIndex=   0
         _Band(0)._MapCol(0)._Alignment=   7
         _Band(0)._MapCol(1)._Name=   "DateHandedIn"
         _Band(0)._MapCol(1)._RSIndex=   1
         _Band(0)._MapCol(2)._Name=   "PTTime"
         _Band(0)._MapCol(2)._RSIndex=   2
         _Band(0)._MapCol(3)._Name=   "PurchaseAmount"
         _Band(0)._MapCol(3)._RSIndex=   3
         _Band(0)._MapCol(3)._Alignment=   7
         _Band(0)._MapCol(4)._Name=   "ExpiryDate"
         _Band(0)._MapCol(4)._RSIndex=   4
         _Band(1).BandIndent=   1
         _Band(1).Cols   =   4
         _Band(1).GridLinesBand=   1
         _Band(1).TextStyleBand=   0
         _Band(1).TextStyleHeader=   0
         _Band(1)._ParentBand=   0
         _Band(1)._NumMapCols=   4
         _Band(1)._MapCol(0)._Name=   "PnCID"
         _Band(1)._MapCol(0)._RSIndex=   0
         _Band(1)._MapCol(0)._Alignment=   7
         _Band(1)._MapCol(1)._Name=   "Quantity"
         _Band(1)._MapCol(1)._RSIndex=   1
         _Band(1)._MapCol(1)._Alignment=   7
         _Band(1)._MapCol(2)._Name=   "Description"
         _Band(1)._MapCol(2)._RSIndex=   2
         _Band(1)._MapCol(3)._Name=   "SerialNr"
         _Band(1)._MapCol(3)._RSIndex=   3
      End
      Begin MSDataGridLib.DataGrid dgCustomerList 
         Height          =   2175
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Visible         =   0   'False
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   3836
         _Version        =   393216
         BackColor       =   14737632
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Customer List"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblCustomerList 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Click in the gray left margin of the grid below to select a customer"
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   1680
         Visible         =   0   'False
         Width           =   6375
      End
      Begin VB.Label lblNoTransactions 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No pawning transactions available for this customer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1440
         TabIndex        =   19
         Top             =   4680
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label lblPNCs 
         BackStyle       =   0  'Transparent
         Caption         =   "Click on any transaction to have it appear to the left"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   4320
         Visible         =   0   'False
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdPawnIn 
      Caption         =   "Pawn In"
      Height          =   550
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8085
      Width           =   1575
   End
   Begin VB.TextBox txtActualBuyBackAmount 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   4560
      TabIndex        =   1
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton cmdGetDetails 
      Caption         =   "Get Details"
      Default         =   -1  'True
      Height          =   550
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtPNCID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Cash&BuyBack ID"
      Top             =   900
      Width           =   735
   End
   Begin MSComCtl2.DTPicker dtpBuyBackDate 
      Height          =   360
      Left            =   840
      TabIndex        =   4
      Top             =   285
      Width           =   2775
      _ExtentX        =   4895
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
      CurrentDate     =   37059
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Buy Back Amount:"
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
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   8205
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CBB ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Cash&BuyBack ID"
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmPawningIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdGetDetails_Click()
    ' Check if PNC has been given
    If Trim(txtPNCID.Text) = "" Then
        MsgBox "Must specify a CBB-ID.", vbCritical, "Missing information"
        ClearFields
        Exit Sub
    End If
    ' Check that recordset is closed, before specifying parameters
    If dePNC.rscmdAllPawningTransactionInformation.State = adStateOpen Then
        dePNC.rscmdAllPawningTransactionInformation.Close
    End If
    ' Specify command parameter and open
    dePNC.cmdAllPawningTransactionInformation CLng(txtPNCID.Text)
    With dePNC.rscmdAllPawningTransactionInformation
        ' Check if recordset could be populated
        If .EOF Then
            MsgBox "The CBB-ID is not valid", vbInformation, "Invalid PNCID"
            ClearFields
            Exit Sub
        End If
        ' Check if this transaction hasn't already been bought back
        If .Fields("Status").Value = 2 Then
            MsgBox "This pawning transaction has already been bought back.", vbInformation, "Cannot proceed"
            txtPNCID.Text = ""
            Exit Sub
        End If
        ' Populate form controls
        lblCustomerInformation.Caption = _
            "Name: " & .Fields("Name").Value & vbCrLf & _
            "IDNr: " & .Fields("IDNr").Value & vbCrLf & _
            "TelW: " & .Fields("TelW").Value & vbCrLf & _
            "TelH: " & .Fields("TelH").Value & vbCrLf & _
            "Address : " & .Fields("Address").Value
        Frame2.Enabled = True
        lblPawningTransaction.Caption = _
            "Date Handed In: " & .Fields("DateHandedIn").Value & vbCrLf & _
            "Purchase Amount: " & Format(.Fields("PurchaseAmount").Value, "currency") & vbCrLf & _
            "Expiry Date: " & .Fields("ExpiryDate").Value & vbCrLf & _
            "Buy Back Amount: " & Format(.Fields("BuyBackAmount").Value, "currency")
            Frame3.Enabled = True
        txtActualBuyBackAmount.Text = .Fields("BuyBackAmount").Value
        Frame4.Enabled = True
        ' If grace period has expired, notify
        If .Fields("ExpiryDateAfter7WorkDays").Value < Date Then
            MsgBox "The transaction is already past the 7 days grace period.", vbCritical, "Transaction Expired"
        End If
        ' Empty lblGoods, before populating it
        lblGoods.Caption = ""
        Do Until .EOF
            lblGoods.Caption = lblGoods.Caption & _
            .Fields("Quantity").Value & "x " & _
            .Fields("Description").Value & " (" & _
            .Fields("SerialNr").Value & ")" & vbCrLf
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub cmdPawnIn_Click()
    'Ensure that minimum information has been given
    If Trim(txtPNCID.Text) = "" Then
        MissingInfo "a PNCID"
        txtPNCID.SetFocus
        Exit Sub
    ElseIf Trim(txtActualBuyBackAmount.Text) = "" Then
        MissingInfo "an actual buyback amount"
        txtActualBuyBackAmount.SetFocus
        Exit Sub
    End If
    
    txtActualBuyBackAmount = Replace(txtActualBuyBackAmount, ",", "")
    
    'Ensure that given txtpncid is displayed
    If Trim(lblCustomerInformation.Caption) = "" Then
        cmdGetDetails_Click
        MsgBox "First review the information, to ensure correctness.", vbInformation, "Review"
        Exit Sub
    End If
    
    On Error GoTo HandleError
    'Connection must be open to start transaction
    If dePNC.cnPNC.State = adStateClosed Then dePNC.cnPNC.Open
    dePNC.cnPNC.BeginTrans
    
    ' Change Status field in PawningTransactions Table
    Dim cmdLocal As ADODB.Command
    Set cmdLocal = New ADODB.Command
    cmdLocal.ActiveConnection = dePNC.cnPNC
    cmdLocal.CommandType = adCmdText
    cmdLocal.CommandText = "UPDATE PawningTransactions SET PawningTransactions.Status = 2, PawningTransactions.ActualBuyBackAmount = " & txtActualBuyBackAmount.Text & ", PawningTransactions.BuyBackDate = #" & dtpBuyBackDate.Value & "# WHERE [PawningTransactions].[PnCID]= " & txtPNCID.Text
    cmdLocal.Execute
    
    ' Change Status field in PawnStock Table
    cmdLocal.CommandText = "UPDATE PawnStock SET PawnStock.Status = 2 WHERE [PawnStock].[PnCID]=" & txtPNCID.Text
    cmdLocal.Execute
    
    ' Insert into CashTransactions
    dePNC.cmdInsertCashTransaction gstrCR, "Pawning In", dtpBuyBackDate.Value, txtPNCID.Text, 0, 0, 0, 0, "", txtActualBuyBackAmount.Text, 0, 0, 0
    
    dePNC.cnPNC.CommitTrans
    MsgBox "Remember to have the customer sign the original CBB document", vbCritical, "IMPORTANT"
    Unload Me
    Exit Sub

HandleError:
    MsgBox "An error occured.  This may have prevented the information you have just entered to be updated to the database.  To ensure that the database stay up to date, take the following steps:" & vbCrLf & _
        "1. On the Main Menu Window, generate a daily report." & vbCrLf & _
        "2. Try locate in the appropriate section of the report, the information you have just entered." & vbCrLf & _
        "3. If you find it, the information has been updated to the database; if not, you should try to redo the transaction." & vbCrLf & _
        "4. If you get this same error, inform the programmer.", vbCritical, "Unsuccessful database update."
    ErrorReport.WriteLine Date & " | " & Time & " | frmPawningIn\cmdPawnIn_Click | " & Err.Description & "(" & Err.Number & ")"
    dePNC.cnPNC.RollbackTrans
    Unload Me
End Sub

Private Sub cmdSearchForPNC_Click()
    ' You must close the recordset before changing the parameter.
    If dePNC.rscmdSearchPNC.State = adStateOpen Then
        dePNC.rscmdSearchPNC.Close
    End If
    'If txtCustID is empty, cmdSearchPNC generates an error
    If txtCustID.Text = "" Then txtCustID.Text = "0"
    'Execute cmdSearchPNC with all parameters, ie reopen recordset
    dePNC.cmdSearchPNC "%" & txtName.Text & "%", "%" & txtIDNr.Text & "%", CLng(txtCustID.Text)
    'Check if result was found
    If dePNC.rscmdSearchPNC.EOF Then
        MsgBox "No result was found for your search conditions", vbInformation, "No result"
    End If
    With dgCustomerList
        .DataMember = "cmdSearchPNC"
        Set .DataSource = dePNC
        .Visible = True
    End With
    lblCustomerList.Visible = True
End Sub

Private Sub cmdSearchForPNCID_Click()
    Me.Width = 12000
End Sub

Private Sub dgCustomerList_Click()
    ' Hide hfgPNCs before resetting values
    hfgPNCs.Visible = False
    ' You must close the recordset before changing the parameter.
    If dePNC.rscmdSearchPNC2.State = adStateOpen Then
        dePNC.rscmdSearchPNC2.Close
    End If
    dePNC.cmdSearchPNC2 CLng(dePNC.rscmdSearchPNC.Fields("CustID").Value)
    'Check whether selected customer has pawningtransactions
    If dePNC.rscmdSearchPNC2.EOF Then
        'Display No Transactions label + hide lblPNCs
        'hfgPNCs.Clear
        lblNoTransactions.Visible = True
        lblPNCs.Visible = False
    Else
        'Set properties of hfgPNCs and display
        With hfgPNCs
            .DataMember = "cmdSearchPNC2"
            Set .DataSource = dePNC
            .ColHeaderCaption(1, 0) = "CBBID"
            hfgPNCs.Visible = True
            'hfgPNCs.CollapseAll
        End With
        'Display assosiated label
        lblPNCs.Visible = True
    End If
End Sub

Private Sub Form_Load()
    dtpBuyBackDate.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' You must close the recordset before unload, to prevent display of data when it re-loads during same run of application
    If dePNC.rscmdAllPawningTransactionInformation.State = adStateOpen Then
        dePNC.rscmdAllPawningTransactionInformation.Close
    End If
    If dePNC.rscmdSearchPNC2.State = adStateOpen Then
        dePNC.rscmdSearchPNC2.Close
    End If
    If dePNC.rscmdSearchPNC.State = adStateOpen Then
        dePNC.rscmdSearchPNC.Close
    End If
End Sub

Private Sub hfgPNCs_Click()
   'hfgPNCs.Text will return the first col value of the selected row
   txtPNCID.Text = hfgPNCs.Text
   cmdGetDetails_Click
End Sub

Private Sub txtActualBuyBackAmount_KeyPress(KeyAscii As Integer)
    'Ensure that only numbers, backspace or decimal point are entered
    If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 46) Then
        MsgBox "Actual BuyBack Amount must have a numeric value.", vbInformation, "Key invalid"
        KeyAscii = 0
    End If
End Sub

Private Sub txtActualBuyBackAmount_LostFocus()
    SetCurr txtActualBuyBackAmount
End Sub

Private Sub txtActualBuyBackAmount_Validate(Cancel As Boolean)
    If Not IsNumeric(txtActualBuyBackAmount.Text) Then
        MsgBox "Actual buyback amount must be a numeric value.", vbInformation, "Value not valid"
        Cancel = True
    End If
End Sub

Private Sub txtCustID_Validate(Cancel As Boolean)
    'Resume only if either the content is numeric or empty
    If Not (IsNumeric(txtCustID.Text) Or Trim(txtCustID.Text) = "") Then
        MsgBox "CustomerID must be a number", vbInformation, "CustomerID"
        Cancel = True
    End If
End Sub

Private Sub txtPNCID_Change()
    'Ensure that only numbers, backspace or decimal point are entered
    If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 46 Or IsEmpty(KeyAscii)) Then
        MsgBox "CBB-ID must have a numeric value.", vbInformation, "Key invalid"
        KeyAscii = 0
    End If
End Sub

Private Sub ClearFields()
    ' Clear populated fields
    txtPNCID.Text = ""
    lblCustomerInformation.Caption = ""
    lblGoods.Caption = ""
    lblPawningTransaction.Caption = ""
    txtActualBuyBackAmount.Text = ""
    txtPNCID.SetFocus
End Sub
