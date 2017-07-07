VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPawningOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pawning Out"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11325
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPawningOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   11325
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   44
      Top             =   5775
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19923
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   " Articles Pledged "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   4320
      TabIndex        =   36
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtQuantity 
         Height          =   360
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Text            =   "1"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtSerialNr 
         Height          =   360
         Index           =   0
         Left            =   5280
         MaxLength       =   30
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtDescription 
         Height          =   360
         Index           =   0
         Left            =   720
         MaxLength       =   75
         TabIndex        =   4
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtQuantity 
         Height          =   360
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Text            =   "1"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtSerialNr 
         Height          =   360
         Index           =   1
         Left            =   5280
         MaxLength       =   30
         TabIndex        =   8
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtDescription 
         Height          =   360
         Index           =   1
         Left            =   720
         MaxLength       =   75
         TabIndex        =   7
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox txtQuantity 
         Height          =   360
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Text            =   "1"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtSerialNr 
         Height          =   360
         Index           =   2
         Left            =   5280
         MaxLength       =   30
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtDescription 
         Height          =   360
         Index           =   2
         Left            =   720
         MaxLength       =   75
         TabIndex        =   10
         Top             =   1440
         Width           =   4455
      End
      Begin VB.TextBox txtQuantity 
         Height          =   360
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Text            =   "1"
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox txtSerialNr 
         Height          =   360
         Index           =   3
         Left            =   5280
         MaxLength       =   30
         TabIndex        =   14
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtDescription 
         Height          =   360
         Index           =   3
         Left            =   720
         MaxLength       =   75
         TabIndex        =   13
         Top             =   1920
         Width           =   4455
      End
      Begin VB.TextBox txtPurchaseAmount 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1440
         TabIndex        =   15
         Text            =   "0"
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtBuyBackAmount 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1440
         TabIndex        =   16
         Text            =   "0"
         Top             =   3120
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpExpiryDate 
         Height          =   360
         Left            =   3960
         TabIndex        =   18
         Top             =   3120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         Format          =   22740992
         CurrentDate     =   37052
      End
      Begin MSComCtl2.DTPicker dtpDateHandedIn 
         Height          =   360
         Left            =   3960
         TabIndex        =   17
         Top             =   2520
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "dd - mmmm - yy"
         Format          =   22740992
         CurrentDate     =   37052
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   255
         Left            =   720
         TabIndex        =   42
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Nr."
         Height          =   255
         Left            =   5280
         TabIndex        =   41
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Amount:"
         Height          =   495
         Left            =   240
         TabIndex        =   40
         Top             =   2445
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Buy Back Amount:"
         Height          =   495
         Left            =   240
         TabIndex        =   39
         Top             =   3045
         Width           =   1455
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Handed In:"
         Height          =   255
         Left            =   2640
         TabIndex        =   38
         Top             =   2565
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date:"
         Height          =   255
         Left            =   2640
         TabIndex        =   37
         Top             =   3165
         Width           =   1215
      End
   End
   Begin VB.Frame frame1 
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
      Height          =   5415
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtCustomerID 
         Height          =   360
         Left            =   960
         TabIndex        =   0
         ToolTipText     =   "Cash&BuyBack ID"
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton cmdGetDetails 
         Caption         =   "Get Details"
         Default         =   -1  'True
         Height          =   550
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Cash&BuyBack"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearchForCustomer 
         Caption         =   "Search for Customer"
         Height          =   550
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4560
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cust ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         ToolTipText     =   "Cash&BuyBack ID"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblID 
         BackStyle       =   0  'Transparent
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
         Left            =   1320
         TabIndex        =   34
         Top             =   1230
         Width           =   1815
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
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
         Left            =   1320
         TabIndex        =   33
         Top             =   795
         Width           =   3015
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel(h):"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2085
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel(w):"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1650
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   795
         Width           =   975
      End
      Begin VB.Label lblTelW 
         BackStyle       =   0  'Transparent
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
         Left            =   1320
         TabIndex        =   27
         Top             =   1650
         Width           =   1815
      End
      Begin VB.Label lblTelH 
         BackStyle       =   0  'Transparent
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
         Left            =   1320
         TabIndex        =   26
         Top             =   2085
         Width           =   1935
      End
      Begin VB.Label lblAddress 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   1320
         TabIndex        =   25
         Top             =   2520
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   550
      Left            =   8280
      TabIndex        =   20
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdPawnOut 
      Caption         =   "Pawn Out"
      Height          =   550
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   6360
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Articles Pledged:"
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
      Left            =   6840
      TabIndex        =   23
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblPTTime 
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
      Height          =   255
      Left            =   9720
      TabIndex        =   22
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
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
      Left            =   8760
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmPawningOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lngPNC As Long 'PNC for next pawning out
Private strMonthNr As String 'MonthNr for next pawning out

Private Function NextPnC() As Long
    With dePNC.rscmdLastPNC
        .Open
        NextPnC = .Fields("LastPNC").Value + 1
        .Close
    End With
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Public Sub cmdGetDetails_Click()
    If Trim(txtCustomerID.Text) = "" Then
        MsgBox "You must specify a valid CustomerID", vbCritical, "Customer ID"
        txtCustomerID.Text = ""
        txtCustomerID.SetFocus
        Exit Sub
    End If
    PopulateCustomerDetails CLng(txtCustomerID.Text)
End Sub

Private Sub cmdPawnOut_Click()
    On Error GoTo HandleError
    
    ClearStockList
    
    'Prepare PNCID and MonthNr
    lngPNC = NextPnC
    strMonthNr = NextMonthNr(lngPNC - 1)
    'Connection must be open to start transaction
    If dePNC.cnPNC.State = adStateClosed Then dePNC.cnPNC.Open
    If ReadyToPawn Then
        dePNC.cnPNC.BeginTrans
        ' add entry to PawningTransactions table
        dePNC.cmdNewPawningTransaction lngPNC, strMonthNr, CDate(lblPTTime.Caption), CLng(txtCustomerID.Text), CDbl(txtPurchaseAmount.Text), CDate(dtpDateHandedIn.Value), CDbl(txtBuyBackAmount.Text), CDate(dtpExpiryDate.Value), FinalExpiryDate(dtpExpiryDate.Value)
        ' Insert into CashTransactions
        dePNC.cmdInsertCashTransaction gstrCR, "Pawning Out", dtpDateHandedIn.Value, lngPNC, 0, 0, 0, 0, "", -(txtPurchaseAmount.Text), 0, 0, 0
        ' add entries to PawnStock table
        Dim i As Integer
        arrStockSold(iStock, 0) = "Quantity"
        arrStockSold(iStock, 1) = "Description"
        arrStockSold(iStock, 2) = "Serial Nr:"
        iStock = iStock + 1
        For i = 0 To 3
            ' check whether descriptions are given
            If Trim(txtDescription(i).Text) <> "" Then
                dePNC.cmdNewPawnStock lngPNC, txtDescription(i).Text, txtSerialNr(i).Text, CInt(txtQuantity(i).Text)
                arrStockSold(iStock, 0) = (Trim(txtQuantity(i).Text))
                arrStockSold(iStock, 1) = (Trim(txtDescription(i).Text))
                arrStockSold(iStock, 2) = (Trim(txtSerialNr(i).Text))
                iStock = iStock + 1
            End If
        Next i
        dePNC.cnPNC.CommitTrans
        MsgBox "Information updated to database", vbInformation, "Database updated"
    Else
        Exit Sub
    End If
    ' Create reports
    On Error GoTo lblErrorReport
    StatusBar1.Panels(1).Text = "Generating Reports ...."
   
    Template_AgreementTearOff CStr(lngPNC), Trim(lblName), Trim(txtBuyBackAmount.Text), _
                              Format((Trim(dtpExpiryDate.Value)), "DD-MMM-YYYY")
   
    Template_Agreement CStr(lngPNC), Trim(strMonthNr), Trim(lblPTTime.Caption), _
                       Trim(txtCustomerID.Text), Trim(lblName), Trim(lblID), _
                       Trim(lblTelW), Trim(lblTelH), Trim(lblAddress), _
                       Trim(txtPurchaseAmount.Text), Trim(txtBuyBackAmount.Text), _
                       Format((Trim(dtpDateHandedIn.Value)), "DD-MMM-YYYY"), _
                       Format((Trim(dtpExpiryDate.Value)), "DD-MMM-YYYY")
                                                     
                       'FinalExpiryDate(dtpExpiryDate.Value)
    'PawningOutReport
    StatusBar1.Panels(1).Text = ""
    Unload Me
    Exit Sub

HandleError:
    MsgBox "An error occured.  This may have prevented the information you have just entered to be updated to the database.  To ensure that the database stay up to date, take the following steps:" & vbCrLf & _
        "1. On the Main Menu Window, generate a daily report." & vbCrLf & _
        "2. Try locate in the appropriate section of the report, the information you have just entered." & vbCrLf & _
        "3. If you find it, the information has been updated to the database; if not, you should try to redo the transaction." & vbCrLf & _
        "4. If you get this same error, inform the programmer.", vbCritical, "Unsuccessful database update."
    ErrorReport.WriteLine Date & " | " & Time & " | frmPawningOut\cmdPawnOut_Click | " & Err.Description & "(" & Err.Number & ")"
    dePNC.cnPNC.RollbackTrans
    StatusBar1.Panels(1).Text = "Error"
    Unload Me

lblErrorReport:
    MsgBox "An error occured during the print of the report/s" & vbCrLf & _
           "You might need to generate the report manually!", vbCritical + vbOKOnly, "Report Error"
    ErrorReport.WriteLine Date & " | " & Time & " | frmPawningOut\cmdPawnOut_Click | " & Err.Description & "(" & Err.Number & ")"
    StatusBar1.Panels(1).Text = "Error in Generating Report"
    Unload Me
End Sub

Private Sub cmdSearchForCustomer_Click()
    'Prepare form before display
    With frmCustomerManagement
        '.Width = 7900
        '.cmdUseThisCustomer.Left = 120
        .cmdAddCustomer.Enabled = False
        .cmdUpdateCustomerInformation.Enabled = False
        .cmdUseThisCustomer.Visible = True
        .Show vbModal
    End With
End Sub

Private Sub Form_Load()
    'Populate static fields
    lblPTTime.Caption = Format(Time, "Short time")
    dtpDateHandedIn.Value = Date
    dtpExpiryDate.Value = Date
    
End Sub

Private Function NextMonthNr(lngPNC As Long) As String
    'Get latest MonthNr
    Dim strLastMonthNr As String
    dePNC.cmdLastMonthNr lngPNC
    strLastMonthNr = dePNC.rscmdLastMonthNr.Fields("MonthNr").Value
    dePNC.rscmdLastMonthNr.Close
    'Get month portion
    Dim intFS1 As Integer
    Dim intFS2 As Integer
    Dim intMonth As Integer 'month portion of last MonthNr
    intFS1 = InStr(1, strLastMonthNr, "/")
    intFS2 = InStr(intFS1 + 1, strLastMonthNr, "/")
    intMonth = Mid(strLastMonthNr, intFS1 + 1, intFS2 - intFS1 - 1)
    'Determine if still in same month as intMonth
    If intMonth = Month(Date) Then
        ' still the same month
        Dim strDay As String
        strDay = Mid(strLastMonthNr, 1, intFS1 - 1)
        NextMonthNr = strDay + 1 & "/" & Month(Date) & "/" & Year(Date)
    Else
        ' beginning of the next month
        NextMonthNr = "1/" & Month(Date) & "/" & Year(Date)
    End If
End Function

Private Sub PopulateCustomerDetails(lngID As Long)
    dePNC.cmdGetCustomerDetails lngID
    With dePNC.rscmdGetCustomerDetails
        ' check if given CustomerID created a valid recordset
        If .EOF Then
            MsgBox "CustomerID does not exist.", vbCritical, "Invalid CustomerID"
            txtCustomerID.Text = ""
            .Close
            Exit Sub
        End If
        lblName.Caption = .Fields("Name").Value
        lblID.Caption = .Fields("IDnr").Value
        lblTelW.Caption = .Fields("TelW").Value
        lblTelH.Caption = .Fields("TelH").Value
        lblAddress.Caption = .Fields("Address").Value
        .Close
    End With
End Sub

Private Function ReadyToPawn() As Boolean
    If Trim(txtCustomerID.Text) = "" Then
        MissingInfo "a Customer ID"
        txtCustomerID.SetFocus
        ReadyToPawn = False
    ElseIf Trim(txtDescription(0).Text) = "" Then
        MissingInfo "at least one pawn item"
        txtDescription(0).SetFocus
        ReadyToPawn = False
    ElseIf CDbl(txtPurchaseAmount.Text) <= 0 Then
        MissingInfo "a Purchase Price"
        txtPurchaseAmount.SetFocus
        ReadyToPawn = False
    ElseIf CDbl(txtBuyBackAmount.Text) <= 0 Then
        MissingInfo "a Buy Back Price"
        txtBuyBackAmount.SetFocus
        ReadyToPawn = False
    ElseIf dtpDateHandedIn.Value > dtpExpiryDate.Value Then
        MsgBox "Expiry Date must be set a date later that the Date Handed in.", vbInformation, "Date Selection"
    ElseIf CCur(txtPurchaseAmount.Text) > CCur(txtBuyBackAmount.Text) Then
        MsgBox "Purchase Amount cannot exceed Buy Back Amount"
    Else
        ReadyToPawn = True
    End If
    
End Function

Public Function FinalExpiryDate(dtmExpire As Date) As Date
    ' Add 7 days, excluding Sundays
    Dim i As Integer
    For i = 1 To 7 ' From day of transaction to day of available to sell
        dtmExpire = DateAdd("w", 1, dtmExpire)
'        If Weekday(dtmExpire) = 1 Then 'Sunday
'            ' Sunday, therefore ignore increment
'            i = i - 1
'        End If
    Next i
    FinalExpiryDate = dtmExpire
End Function

Private Sub PawningOutReport()
    ' You must close the recordset before changing the parameter.
    If dePNC.rscmdPawningOutReport.State = adStateOpen Then
        dePNC.rscmdPawningOutReport.Close
    End If
    'Execute cmdPawningOutReport with required parameters
    dePNC.cmdPawningOutReport lngPNC
    'Display/Print PawningOutReport
    With drPawningOut
        .DataMember = "cmdPawningOutReport"
        Set .DataSource = dePNC
        .Refresh
    End With
    drPawningOut.Show vbModal
    'drPawningOut.PrintReport
    'Display/Print PawningOutTearOffReport
    With drPawningOutTearOff
        .DataMember = "cmdPawningOutReport"
        Set .DataSource = dePNC
        .Refresh
    End With
    drPawningOutTearOff.Show vbModal
    'drPawningOutTearOff.PrintReport


'**************************************
' Code to use Word to generate reports
'**************************************
'    'Application Path
'    Dim AppPath As String
'    AppPath = App.Path
'
'' Problem: Cannot close Word applications
'' Solution: Create 1 public application object to use throughout application
'    Dim WordApp As Word.Application
'    Set WordApp = New Word.Application
'
'    Dim doc As Word.Document
'    Set doc = WordApp.Documents.Open(AppPath & "\PawningOutReport.dot")
'
'    doc.PrintPreview
'Dim rngRange As Word.Range
'
'Set rngRange = doc.Bookmarks("Test").Range
'rngRange.InsertAfter "Seattle, WA 12345"
'rngRange.Select
'    doc.PrintOut
'    doc.Close wdDoNotSaveChanges
'    WordApp.Quit
'    Set doc = Nothing
'    Set WordApp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Close all remaining open recordsets
    If dePNC.rscmdGetCustomerDetails.State = adStateOpen Then
        dePNC.rscmdGetCustomerDetails.Close
    End If
    If dePNC.rscmdPawningOutReport.State = adStateOpen Then
        dePNC.rscmdPawningOutReport.Close
    End If
End Sub

Private Sub txtBuyBackAmount_KeyPress(KeyAscii As Integer)
    'Ensure that only numbers, backspace or decimal point are entered
    If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 46) Then
        MsgBox "BuyBack Amount must have a numeric value.", vbInformation, "Key invalid"
        KeyAscii = 0
    End If
End Sub

Private Sub txtBuyBackAmount_LostFocus()
    SetCurr txtBuyBackAmount
End Sub

Private Sub txtBuyBackAmount_Validate(Cancel As Boolean)
    If Not IsNumeric(txtBuyBackAmount.Text) Then
        MsgBox "Buy Back Amount must be a numeric value.", vbInformation, "Numeric Value"
        Cancel = True
    End If
End Sub

Private Sub txtCustomerID_KeyPress(KeyAscii As Integer)
    'Ensure that only numbers, backspace or decimal point are entered
    If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 46) Then
        MsgBox "CustomerID must have a numeric value.", vbInformation, "Key invalid"
        KeyAscii = 0
    End If
End Sub

Private Sub txtPurchaseAmount_KeyPress(KeyAscii As Integer)
    'Ensure that only numbers, backspace or decimal point are entered
    If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 46) Then
        MsgBox "Purchase Amount must have a numeric value.", vbInformation, "Key invalid"
        KeyAscii = 0
    End If
End Sub

Private Sub txtPurchaseAmount_LostFocus()
    SetCurr txtPurchaseAmount
End Sub

Private Sub txtPurchaseAmount_Validate(Cancel As Boolean)
    If Not IsNumeric(txtPurchaseAmount.Text) Then
        MsgBox "Purchase Amount must be a numeric value.", vbInformation, "Numeric Value"
        Cancel = True
    End If
End Sub

Private Sub txtQuantity_Validate(Index As Integer, Cancel As Boolean)
    If Not IsNumeric(txtQuantity(Index).Text) Then
        MsgBox "Quantity must be a numeric value and less than 255.", vbInformation, "Numeric Value"
        Cancel = True
    End If
End Sub
