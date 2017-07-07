VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   9270
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9585
   FillColor       =   &H80000004&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   9585
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      Caption         =   " Stock Code List: "
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
      Height          =   1815
      Left            =   240
      TabIndex        =   37
      Top             =   6840
      Width           =   4455
      Begin VB.CommandButton cmdStandardStockCodeList 
         Caption         =   "Standard Stock"
         Height          =   550
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton cmdTangoStockCodeLsit 
         Caption         =   "Tango Stock"
         Height          =   550
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1080
         Width           =   3255
      End
   End
   Begin VB.ComboBox cboCashRegister 
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
      Left            =   6960
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   30
      Top             =   8850
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
            Text            =   "Cash Register: Unknown"
            TextSave        =   "Cash Register: Unknown"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "13:11"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   550
      Left            =   5880
      TabIndex        =   19
      Top             =   7475
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   " Pawning History: "
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
      Height          =   2175
      Left            =   4800
      TabIndex        =   26
      Top             =   2520
      Width           =   4455
      Begin VB.CommandButton cmdPawningHistory 
         Caption         =   "Pawning History"
         Height          =   550
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1320
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   360
         Left            =   1320
         TabIndex        =   12
         Top             =   840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   147521536
         CurrentDate     =   37073
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   360
         Left            =   1320
         TabIndex        =   11
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   147521536
         CurrentDate     =   37073
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "End Date:"
         Height          =   225
         Left            =   240
         TabIndex        =   28
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   420
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Report: "
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
      Height          =   2175
      Left            =   240
      TabIndex        =   25
      Top             =   2520
      Width           =   4455
      Begin VB.CommandButton cmdCustomReport 
         Caption         =   "Custom Report"
         Height          =   550
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Report for period Start - End date"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdDailyReport 
         Caption         =   "Daily Report"
         Height          =   550
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Only running report for 'End Date'"
         Top             =   1320
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   360
         Left            =   1200
         TabIndex        =   9
         Top             =   840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   130416640
         CurrentDate     =   37073
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   360
         Left            =   1200
         TabIndex        =   31
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   130416640
         CurrentDate     =   37073
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "End Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   900
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "  Stock: "
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
      Height          =   1815
      Left            =   4800
      TabIndex        =   24
      Top             =   4920
      Width           =   4455
      Begin VB.CommandButton cmdReliaze_Income 
         Caption         =   "Realizable Income"
         Height          =   550
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdMonthlyReportNormalStock 
         Caption         =   "Stock Report"
         Height          =   550
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Normal Standard & Tango Stock Report "
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdMonthlyReportPawnStock 
         Caption         =   "Pawn Report"
         Height          =   550
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdStockSearch 
         Caption         =   "Search"
         Height          =   550
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   " Cash Transactions: "
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
      Height          =   2295
      Left            =   6480
      TabIndex        =   23
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton cmdOtherExpenses 
         Caption         =   "Other Expenses"
         Height          =   550
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CommandButton cmdStockSale 
         Caption         =   "Sale (Pawn and Stock)"
         Height          =   550
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton cmdStockPurchase 
         Caption         =   "Stock Purchase"
         Height          =   550
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Stock Management: "
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
      Height          =   1815
      Left            =   240
      TabIndex        =   22
      Top             =   4920
      Width           =   4455
      Begin VB.CommandButton cmdTangoStock 
         Caption         =   "Tango Stock"
         Height          =   550
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CommandButton cmdCellphoneStock 
         Caption         =   "Standard Stock"
         Height          =   550
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton cmdLayBuyPaymentOrCancellation 
         Caption         =   "Payment or Cancellation"
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdNewLayBuy 
         Caption         =   "New"
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Pawning: "
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
      Height          =   2295
      Left            =   3360
      TabIndex        =   21
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton cmdManageExpiredPawnings 
         Caption         =   "Expired"
         Height          =   550
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CommandButton cmdPawningIn 
         Caption         =   "In"
         Height          =   550
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton cmdPawningOut 
         Caption         =   "Out"
         Height          =   550
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Customer Maintenance: "
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
      Height          =   2295
      Left            =   240
      TabIndex        =   20
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton cmdUpdateCustomer 
         Caption         =   "Update"
         Height          =   550
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CommandButton cmdSearchCustomer 
         Caption         =   "Search"
         Height          =   550
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton cmdAddCustomer 
         Caption         =   "Add"
         Height          =   550
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Register:"
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
      Left            =   5760
      TabIndex        =   34
      Top             =   8280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuSettingsDatabase 
         Caption         =   "Database Location"
         Begin VB.Menu mnuSettingsDatabaseView 
            Caption         =   "View"
         End
         Begin VB.Menu mnuSettingsDatabaseSet 
            Caption         =   "Set or Change"
         End
      End
      Begin VB.Menu mnuSettingsCashRegister 
         Caption         =   "Cash Register Location"
      End
      Begin VB.Menu mnuCurrency 
         Caption         =   "Currency"
         Visible         =   0   'False
         Begin VB.Menu mnuCurrencyView 
            Caption         =   "View"
         End
         Begin VB.Menu mnuCurrencySet 
            Caption         =   "Set or Change"
         End
      End
      Begin VB.Menu mnuAddress 
         Caption         =   "Company Address"
         Visible         =   0   'False
         Begin VB.Menu mnuAddressView 
            Caption         =   "View"
         End
         Begin VB.Menu mnuAddressSet 
            Caption         =   "Set or Change"
         End
      End
      Begin VB.Menu mnuVat 
         Caption         =   "Vat Number"
         Visible         =   0   'False
         Begin VB.Menu mnuVatView 
            Caption         =   "View"
         End
         Begin VB.Menu mnuVatSet 
            Caption         =   "Set or Update"
         End
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMaintenance 
      Caption         =   "&Maintenance"
      Begin VB.Menu mnuDBMaintenance 
         Caption         =   "Database Maintenance"
         Begin VB.Menu mnuDBCTHistory 
            Caption         =   "Clear History"
         End
         Begin VB.Menu mnuDBTango 
            Caption         =   "Change Tango Entries"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboCashRegister_Change()
StatusBar_One gstrCR

End Sub

Private Sub cmdAddCustomer_Click()
    'prepare form before shown
    With frmCustomerManagement
        .fraNewCustomer.Left = 120
        .fraNewCustomer.Top = 3840
        '.Width = 3500
        '.txtCustID.Visible = False
        '.lblCustID.Visible = False
        '.cmdClearFields.Visible = False
        .txtCustID.Enabled = False
        .cmdUpdateCustomerInformation.Enabled = False
        .Show vbModal
    End With
End Sub

Private Sub cmdCellphoneStock_Click()
    frmCellphoneManagement.Show vbModal
End Sub

Private Sub cmdCustomReport_Click()
Dim iDateDiff As Integer
   
    iDateDiff = DateDiff("d", dtpDateFrom.Value, dtpDate.Value)
    If iDateDiff < 0 Then
        MsgBox "'Start Date' can not be later than 'End Date'", vbCritical + vbOKOnly, "Daily Report Error"
        Exit Sub
    End If
    
    If dePNC.rscmdCustomReport_Grouping.State = adStateOpen Then
        dePNC.rscmdCustomReport_Grouping.Close
    End If
    
    dePNC.cmdCustomReport_Grouping dtpDateFrom.Value, dtpDate.Value
    
    With drCustomReport
        .DataMember = "cmdCustomReport_Grouping"
        Set .DataSource = dePNC
        .Sections("PageHeader").Controls("Label18").Caption = "Custom Report for the period: " & dtpDateFrom.Value & " - " & dtpDate.Value
        .Refresh
    End With
    drCustomReport.Show vbModal

End Sub

Private Sub cmdDailyReport_Click()
    If dePNC.rscmdDailyReport_Grouping.State = adStateOpen Then
        dePNC.rscmdDailyReport_Grouping.Close
    End If
         
    'Execute cmdDailyReport_Grouping with required parameters
    dePNC.cmdDailyReport_Grouping dtpDate.Value, cboCashRegister.Text
    
    'Display/Print cmdDailyReport_Grouping
    With drDailyReport
        .DataMember = "cmdDailyReport_Grouping"
        Set .DataSource = dePNC
        .Refresh
    End With
    drDailyReport.Show vbModal

End Sub

Private Sub OterReport()
Dim dtValue As Date
Dim iDateDiff As Integer
Dim i As Integer
       
    iDateDiff = DateDiff("d", dtpDateFrom.Value, dtpDate.Value)
    If iDateDiff < 0 Then
        MsgBox "'To date' can not be smaller as 'From Date'", vbCritical + vbOKOnly, "Daily Report Error"
        Exit Sub
    End If
    
    If iDateDiff > 7 Then
        MsgBox "Maximum of only 7 days allowed! ", vbCritical + vbOKOnly, "Daily Report Error"
        Exit Sub
    End If
    
    dtValue = dtpDateFrom.Value
    
    For i = 0 To iDateDiff
       ' MsgBox dtValue & " - " & iDateDiff
        ' You must close the recordset before changing the parameter.
        If dePNC.rscmdDailyReport_Grouping.State = adStateOpen Then
            dePNC.rscmdDailyReport_Grouping.Close
        End If
        
        'Execute cmdDailyReport_Grouping with required parameters
        'dePNC.cmdDailyReport_Grouping dtpDate.Value, cboCashRegister.Text
        dePNC.cmdDailyReport_Grouping dtValue, cboCashRegister.Text
        
        'Display/Print cmdDailyReport_Grouping
        With drDailyReport
            .DataMember = "cmdDailyReport_Grouping"
            Set .DataSource = dePNC
            .Refresh
        End With
        drDailyReport.Show vbModal
        dtValue = DateAdd("d", 1, dtValue)
    Next i

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLayBuyPaymentOrCancellation_Click()
    frmLayBuyPayment.Show vbModal
End Sub

Private Sub cmdManageExpiredPawnings_Click()
    Load frmExpiredPawnTransactions
    frmExpiredPawnTransactions.LoadDecision
End Sub

Private Sub cmdMonthlyReportNormalStock_Click()
    drMonthlyReportNormalStock.Show vbModal
End Sub

Private Sub cmdMonthlyReportPawnStock_Click()
    drMonthlyReportPawnStock.Show vbModal
End Sub

Private Sub cmdNewLayBuy_Click()
    frmLayBuy.Show vbModal
End Sub

Private Sub cmdOtherExpenses_Click()
    frmOtherExpenses.Show vbModal
End Sub

Private Sub cmdPawningHistory_Click()
    ' You must close the recordset before changing the parameter.
    If dePNC.rscmdPawningHistory.State = adStateOpen Then
        dePNC.rscmdPawningHistory.Close
    End If
    'Execute cmdPawningHistory with required parameters
    dePNC.cmdPawningHistory dtpStartDate.Value, dtpEndDate.Value
    'Display/Print cmdPawningHistory
    With drPawningHistory
        .DataMember = "cmdPawningHistory"
        Set .DataSource = dePNC
        .Refresh
    End With
    drPawningHistory.Show vbModal
End Sub

Private Sub cmdPawningIn_Click()
    'frmPawningIn.Width = 4500
    frmPawningIn.Show vbModal
End Sub

Private Sub cmdPawningOut_Click()
    frmPawningOut.Show vbModal
End Sub

Private Sub cmdReliaze_Income_Click()
    drReportRealizableIncome.Show vbModal

End Sub

Private Sub cmdSearchCustomer_Click()
    'Prepare form before display
    With frmCustomerManagement
        '.Width = 7900
        .cmdAddCustomer.Enabled = False
        .cmdUpdateCustomerInformation.Enabled = False
        .cmdShowHistory.Visible = True
        .cmdShowHistory.Left = 1902
        .Show vbModal
    End With
End Sub

Private Sub cmdStandardStockCodeList_Click()
    drStandardStockCodeList.Show vbModal
End Sub

Private Sub cmdStockPurchase_Click()
    frmStockPurchase.Show vbModal
End Sub

Private Sub cmdStockSale_Click()
    frmStockSale.Show vbModal
End Sub

Private Sub cmdStockSearch_Click()
    frmStockSearch.Show vbModal
End Sub

Private Sub cmdTangoStock_Click()
    frmTangoManagement.Show vbModal
End Sub

Private Sub cmdTangoStockCodeLsit_Click()
    drTangoStockCodeList.Show vbModal
End Sub

Private Sub cmdUpdateCustomer_Click()
    'show update customer button,
    'show Customer Management form and
    'hide update customer button
    With frmCustomerManagement
        '.fraCustomerUpdate.Left = 120
        '.fraCustomerUpdate.Top = 4440
        '.Width = 7900
        .cmdAddCustomer.Enabled = False
        .Show vbModal
    End With
End Sub

Private Sub Form_Load()
    dtpDate.Value = Date
    dtpDateFrom.Value = Date
    dtpEndDate.Value = Date
    dtpStartDate.Value = DateAdd("m", -1, Date)
    
    'populate cboCashRegister
    rs.Source = "Select distinct cashregister from cashtransactions"
    rs.Open
    Do Until rs.EOF
        If IsNull(rs.Fields("cashregister").Value) Then
            cboCashRegister.AddItem ""
            frmCashRegister.cboCashRegister.AddItem ""
        Else
            cboCashRegister.AddItem UCase(rs.Fields("cashregister").Value)
            frmCashRegister.cboCashRegister.AddItem UCase(rs.Fields("cashregister").Value)
        End If
        rs.MoveNext
    Loop
    'Ensure that recordset is closed
    If rs.State = adStateOpen Then
        rs.Close
    End If
    'Make registry setting for CashRegister the default selected
    CashRegisterExists gstrCR
    StatusBar_One gstrCR
    
    GetShop

    
    frmMain.Caption = "Cash & BuyBack - v. " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Function GetShop()
    
    sTemplateShop = Trim(Command())
    
    If sTemplateShop <> "" Then
        sTemplateShop = UCase(sTemplateShop)
        StatusBar_Two sTemplateShop
        If ValidTemplates("Agreement.doc") <> True Then
            'TemplateError
            Exit Function
        End If
    Else
        MsgBox "Please provide 'Shop name' as startup parameter!", vbCritical + vbOKOnly, App.EXEName
        Unload Me
    End If
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    'If any thing else is still running ...
    End
End Sub

Private Sub mnuAddressSet_Click()
    Dim strPW As String
    strPW = InputBox("Password is required to change setting." & vbCrLf & "Please specify.", "Setting Password")
    If UCase(strPW) <> "7777" Then
        MsgBox "Illegal password", vbInformation, "Password Incorrect"
        Exit Sub
    End If
    
    Load frmAddress
    frmAddress.Show vbModal
    
End Sub

Private Sub mnuAddressView_Click()
    MsgBox "The Address is set to:" & vbCrLf & gstrAddress1 & vbCrLf & gstrAddress2 & vbCrLf & gstrAddress3, vbInformation, "Address"
    
End Sub

Private Sub mnuCurrencySet_Click()
    Dim strPW As String
    strPW = InputBox("Password is required to change setting." & vbCrLf & "Please specify.", "Setting Password")
    If UCase(strPW) <> "7777" Then
        MsgBox "Illegal password", vbInformation, "Password Incorrect"
        Exit Sub
    End If
    
    Dim strResp As String
    strResp = SetCurrency
    If strResp = "not set" Then
        MsgBox "The currency has not been changed.", vbInformation, "Currency Symbol Unchanged"
    Else
        MsgBox "The currency has been successfully set to " & strResp, vbInformation, "Setting Successfull"
        gstrCR = strResp
    End If

End Sub

Private Sub mnuCurrencyView_Click()
    MsgBox "The Currency is set to:" & vbCrLf & gstrCurrency, vbInformation, "Currency"
End Sub

Private Sub mnuDBCTHistory_Click()
    Dim strPW As String
    strPW = InputBox("Password is required to change setting." & vbCrLf & "Please specify.", "Setting Password")
    If UCase(strPW) <> "7777" Then
        MsgBox "Illegal password", vbInformation, "Password Incorrect"
        Exit Sub
    End If

    Load frmDBMaintenance
    frmDBMaintenance.Show vbModal

End Sub

Private Sub mnuDBTango_Click()
    Dim strPW As String
    strPW = InputBox("Password is required to change setting." & vbCrLf & "Please specify.", "Setting Password")
    If UCase(strPW) <> "7777" Then
        MsgBox "Illegal password", vbInformation, "Password Incorrect"
        Exit Sub
    End If
    
    Load frmDBTango
    frmDBTango.Show vbModal
    
End Sub

Private Sub mnuExit_Click()
cmdExit_Click

End Sub

'Private Sub mnuSettingsCashRegisterSet_Click()
'    'Check for valid password
'    Dim strPW As String
'    strPW = InputBox("Password is required to change setting." & vbCrLf & "Please specify.", "Setting Password")
'    If UCase(strPW) <> "7777" Then
'        MsgBox "Illegal password", vbInformation, "Password Incorrect"
'        Exit Sub
'    End If
'
'    Dim strResp As String
'    strResp = SetCashRegister
'    If strResp = "not set" Then
'        MsgBox "The cash register location has not been changed.", vbInformation, "Cash Register Unchanged"
'    Else
'        MsgBox "The cash register location has been successfully set to " & strResp, vbInformation, "Setting Successfull"
'        gstrCR = strResp
'        AddCashRegister strResp
'    End If
'End Sub
'
'Private Sub mnuSettingsCashRegisterView_Click()
'Dim sNewCR As String
'    Dim strPW As String
'    strPW = InputBox("Password is required to change setting." & vbCrLf & "Please specify.", "Setting Password")
'    If UCase(strPW) <> "7777" Then
'        MsgBox "Illegal password", vbInformation, "Password Incorrect"
'        Exit Sub
'    End If
'
'    'MsgBox "The Cash Register Location is set to:" & vbCrLf & gstrCR, vbInformation, "Cash Register Location"
'    frmCashRegister.lblCR.Caption = gstrCR
'    frmCashRegister.strCashRegiter = gstrCR
'    frmCashRegister.LoadCR
'    Load frmCashRegister
'    frmCashRegister.Show vbModal
'    sNewCR = frmCashRegister.strCashRegiter
'    'Unload frmCashRegister
'
'    If sNewCR <> "" Then StatusBar_One sNewCR
'
'End Sub

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

Public Function SetCurrency() As String
    Dim strCR As String
    strCR = UCase(InputBox("Type the currency symbol. (e.g. $)", "Currency"))
    If Len(strCR) > 4 Then
        MsgBox "The currency symbol may not exceed 4 characters", vbInformation, "Currency too long"
        SetCurrency = "not set"
        Exit Function
    End If
    If Len(Trim(strCR)) <> 0 Then
        SaveSetting "Cash&BuyBack", "Settings", "Currency", strCR
        gstrCurrency = strCR
        SetCurrency = strCR
    Else
        SetCurrency = "not set"
    End If
End Function

Public Function SetVAT() As String
    Dim strCR As String
    strCR = UCase(InputBox("Type the VAT Number.", "VAT Number"))
    If Len(strCR) = 0 Then
        MsgBox "The VAT Number may not be 0 length", vbInformation, "Invalid VAT Number"
        SetVAT = "not set"
        Exit Function
    End If
    If Len(Trim(strCR)) <> 0 Then
        SaveSetting "Cash&BuyBack", "Settings", "VatNumber", strCR
        gstrVat = strCR
        SetVAT = strCR
    Else
        SetVAT = "not set"
    End If
End Function

Private Sub mnuSettingsCashRegister_Click()
Dim sNewCR As String
    Dim strPW As String
    strPW = InputBox("Password is required to change setting." & vbCrLf & "Please specify.", "Setting Password")
    If UCase(strPW) <> "7777" Then
        MsgBox "Illegal password", vbInformation, "Password Incorrect"
        Exit Sub
    End If
    
    'MsgBox "The Cash Register Location is set to:" & vbCrLf & gstrCR, vbInformation, "Cash Register Location"
    frmCashRegister.lblCR.Caption = gstrCR
    frmCashRegister.strCashRegiter = gstrCR
    frmCashRegister.LoadCR
    Load frmCashRegister
    frmCashRegister.Show vbModal
    sNewCR = frmCashRegister.strCashRegiter
    'Unload frmCashRegister
    
    If sNewCR <> "" Then StatusBar_One sNewCR
    
End Sub

Private Sub mnuSettingsDatabaseSet_Click()
    'Check for valid password
    Dim strPW As String
    strPW = InputBox("Password is required to change setting." & vbCrLf & "Please specify.", "Setting Password")
    If UCase(strPW) <> "7777" Then
        MsgBox "Illegal password", vbInformation, "Password Incorrect"
        Exit Sub
    End If
    
    Dim strResp As String
    strResp = frmUtilities.GetDBPath
    If strResp = "not set" Then
        MsgBox "The database location has not been changed.", vbInformation, "Database Unchanged"
    Else
        MsgBox "The path to the database has been successfully set to " & strResp, vbInformation, "Setting Successfull"
        MsgBox "Restart the application to activate the new setting.", vbInformation, "Application Shutdown"
        End
    End If
End Sub

Private Sub mnuSettingsDatabaseView_Click()
    MsgBox "Database Location is currently set to:" & vbCrLf & DBPath, vbInformation, "Database Location"
End Sub

Public Sub AddCashRegister(strCR As String)
    If CashRegisterExists(strCR) <> True Then
        cboCashRegister.AddItem strCR
        'Make item selected in cbo
        cboCashRegister.ListIndex = cboCashRegister.NewIndex
    End If
End Sub
    
'Determine if CashRegister already exists in cbo
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

Private Sub mnuVatSet_Click()
    Dim strPW As String
    strPW = InputBox("Password is required to change setting." & vbCrLf & "Please specify.", "Setting Password")
    If UCase(strPW) <> "7777" Then
        MsgBox "Illegal password", vbInformation, "Password Incorrect"
        Exit Sub
    End If
    
    Dim strResp As String
    strResp = SetVAT
    If strResp = "not set" Then
        MsgBox "The VAT Number has not been changed.", vbInformation, "VAT Number Unchanged"
    Else
        MsgBox "The VAT Number has been successfully set to " & strResp, vbInformation, "Setting Successfull"
        gstrVat = strResp
    End If

End Sub

Private Sub mnuVatView_Click()
    MsgBox "VAT Number is currently set to:" & vbCrLf & gstrVat, vbInformation, "VAT Number"
End Sub

Private Function StatusBar_One(Message As String)
StatusBar1.Panels(1).Text = "Cash Register: " & Message
StatusBar1.Refresh

End Function

Private Function StatusBar_Two(Message As String)
StatusBar1.Panels(2).Text = "Shop: " & Message
StatusBar1.Refresh

End Function

