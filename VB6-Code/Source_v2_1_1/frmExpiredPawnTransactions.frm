VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmExpiredPawnTransactions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expired Pawning Transactions"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExpiredPawnTransactions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   8445
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "E&xit"
      Height          =   555
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdReportAndMove 
      Caption         =   "Generate Report and Move Stock"
      Height          =   555
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton cmdClose2 
      Caption         =   "E&xit"
      Height          =   550
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc adoExpiredPawnTransactions 
      Height          =   375
      Left            =   480
      Top             =   7080
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoExpiredPawnTransactions"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoExpiredPawnStock 
      Height          =   375
      Left            =   5760
      Top             =   7080
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoExpiredPawnStock"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dgExpiredPawningStock 
      Bindings        =   "frmExpiredPawnTransactions.frx":08CA
      Height          =   1335
      Left            =   480
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   2355
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   14737632
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
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
      Caption         =   "Items from Select Transaction"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Quantity"
         Caption         =   "Qty"
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
         DataField       =   "Description"
         Caption         =   "Description"
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
      BeginProperty Column02 
         DataField       =   "SerialNr"
         Caption         =   "SerialNr"
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
      BeginProperty Column03 
         DataField       =   "Price"
         Caption         =   "Selling Price"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """N$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "PnCID"
         Caption         =   "PnCID"
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
      BeginProperty Column05 
         DataField       =   "Status"
         Caption         =   "Status"
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
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4004.788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgExpiredPawningTransactions 
      Bindings        =   "frmExpiredPawnTransactions.frx":08EC
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   14737632
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
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
      Caption         =   "Expired Pawning Transactions"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Name"
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
         DataField       =   "DateHandedIn"
         Caption         =   "DateHandedIn"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "d-MMM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "ExpiryDateAfter7Workdays"
         Caption         =   "ExpiryDateAfter7Workdays"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "d-MMM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "PurchaseAmount"
         Caption         =   "PurchaseAmount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """N$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "BuybackAmount"
         Caption         =   "BuybackAmount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """N$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PnCID"
         Caption         =   "PnCID"
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
      BeginProperty Column06 
         DataField       =   "Status"
         Caption         =   "Status"
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2129.953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin VB.Label lblExpiredPawningStock 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Specify a Selling Price for each of the items."
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmExpiredPawnTransactions.frx":0915
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmExpiredPawnTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose2_Click()
    Unload Me
End Sub


Private Sub cmdReportAndMove_Click()
Dim lastCmd As String

    'Confirm that recordset has records
    If adoExpiredPawnTransactions.Recordset.EOF Then
        Exit Sub
    End If
    
    ' Assure that changed field values have been updated
    With adoExpiredPawnStock
        If .RecordSource <> "" Then
            If Not .Recordset.EOF Then
                .Recordset.Update
                .RecordSource = "SELECT Quantity, Description, SerialNr, Price, PnCID, Status FROM PawnStock where PnCID =" & CLng(adoExpiredPawnTransactions.Recordset.Fields("PnCID").Value)
                lastCmd = "A"
                .Refresh
            End If
        End If
    End With
    
    ' Print report for expired pawn transaction goods
    ' to be moved to general sales area
    ' You must close the recordset before changing the parameter.
    If dePNC.rscmdPawnStockToFloor.State = adStateOpen Then
        dePNC.rscmdPawnStockToFloor.Close
        lastCmd = "B"
    End If
    'Execute cmdPawnStockToFloor with required parameter
    dePNC.cmdPawnStockToFloor Date
    'Display/Print PawningOutReport
    With drPawnStockToFloor
        .DataMember = "cmdPawnStockToFloor"
        Set .DataSource = dePNC
        lastCmd = "C"
        .Refresh
    End With
    drPawnStockToFloor.Show vbModal
    'drPawnStockToFloor.PrintReport
    
    '**'Set Status Values of PawningTransactions to 3 (=Expired)
    With adoExpiredPawnTransactions.Recordset
        .MoveFirst 'Start at the beginning
        lastCmd = "D"
        On Error GoTo ErrorTrap
        Do Until .EOF
NextPnC:
            cn.BeginTrans
            .Fields("Status") = 3
            '**'Set Status Values of PawningStock to 3 (=Expired)
            PawnStockToFloor CLng(.Fields("PnCID").Value)
            lastCmd = "E"
            cn.CommitTrans
            .MoveNext
        Loop
        .Close
    End With
    Unload Me
    Exit Sub
    
ErrorTrap:
    MsgBox "An error occured.  This may have prevented the information you have just entered to be updated to the database.  To ensure that the database stay up to date, take the following steps:" & vbCrLf & _
        "1. On the Main Menu Window, generate a daily report." & vbCrLf & _
        "2. Try locate in the appropriate section of the report, the information you have just entered." & vbCrLf & _
        "3. If you find it, the information has been updated to the database; if not, you should try to redo the transaction." & vbCrLf & _
        "4. If you get this same error, inform the programmer.", vbCritical, "Unsuccessful database update."
    ErrorReport.WriteLine Date & " | " & Time & " | frmExpiredPawnTransactions\cmdReportAndMove_Click | " & Err.Description & "(" & Err.Number & ")"
    cn.RollbackTrans
    'Skip current record
    Resume NextPnC
End Sub

Private Sub dgExpiredPawningStock_LostFocus()
    ' Ensure that all changes have been updated
    adoExpiredPawnStock.Recordset.Update
End Sub

Private Sub dgExpiredPawningTransactions_Click()
    With adoExpiredPawnStock
        If adoExpiredPawnStock.RecordSource <> "" Then
            ' Assure that changed field values have been updated
            If Not .Recordset.EOF Then .Recordset.Update
        End If
        .RecordSource = "SELECT Quantity, Description, SerialNr, Price, PnCID, Status FROM PawnStock where PnCID =" & CLng(adoExpiredPawnTransactions.Recordset.Fields("PnCID").Value)
        .Refresh
    End With
    dgExpiredPawningStock.Visible = True
    lblExpiredPawningStock.Visible = True
    cmdReportAndMove.Visible = True
    cmdClose.Visible = False
    cmdClose2.Visible = True
End Sub

Private Sub dgExpiredPawningTransactions_GotFocus()
    'Refresh ADODCs to ensure full population
    adoExpiredPawnTransactions.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Close all remaining open recordsets
    If dePNC.rscmdPawnStockToFloor.State = adStateOpen Then
        dePNC.rscmdPawnStockToFloor.Close
    End If
    If dePNC.rscmdSearchPNC2.State = adStateOpen Then
        dePNC.rscmdSearchPNC2.Close
    End If
End Sub

'Updates status value (to 3) of pawn items for a specific pncid
Private Sub PawnStockToFloor(pncid As Long)
    ' Assure that changed field values have been updated
    ' Populate recordset
    With adoExpiredPawnStock
        If Not .Recordset.EOF Then .Recordset.Update
        .Recordset.Close
        .RecordSource = "SELECT Quantity, Description, SerialNr, Price, PnCID, Status FROM PawnStock where PnCID =" & pncid
        .Refresh
    End With
    ' Update Status values to 3 (=Floor)
    With adoExpiredPawnStock.Recordset
        'On Error Resume Next
        Do Until .EOF
            .Fields("Status") = 3
            .MoveNext
        Loop
    End With
End Sub

Public Sub LoadDecision()
Dim dbDate As Date
    dbDate = CDate("03/01/2016")
    
    ' Set connectionstring for adoExpiredPawnStock
    adoExpiredPawnStock.ConnectionString = cn '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PAWNnCASH.mdb;Persist Security Info=False"
    adoExpiredPawnTransactions.ConnectionString = cn '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PAWNnCASH.mdb;Persist Security Info=False"
    adoExpiredPawnTransactions.RecordSource = "SELECT Customers.Name, PawningTransactions.DateHandedIn, PawningTransactions.ExpiryDateAfter7WorkDays, PawningTransactions.PurchaseAmount, PawningTransactions.BuybackAmount, PawningTransactions.PnCID, PawningTransactions.Status FROM Customers INNER JOIN PawningTransactions ON Customers.CustID = PawningTransactions.CustID WHERE (PawningTransactions.ExpiryDateAfter7WorkDays <= #" & dbDate & "#) AND (PawningTransactions.Status = 1) ORDER BY Customers.Name, PawningTransactions.DateHandedIn"
    adoExpiredPawnTransactions.Refresh
    ' If recordset is empty, display message and unload
    If adoExpiredPawnTransactions.Recordset.EOF Then
        MsgBox "No pawning transactions have expired, since you last checked", vbInformation, "No expired transactions"
        Unload Me
    Else
        Me.Show vbModal
    End If
End Sub
