VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmStockSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search for Available Stock"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStockSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7965
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdSellThisItem 
      Caption         =   "Sell this Item"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame fraStockSell 
      Caption         =   "To Select a Stock Item to Sell:"
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Stock Item by clicking in the margin next to the item and click on Sell this Item."
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
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   150
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid dgStock 
      Bindings        =   "frmStockSearch.frx":08CA
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5741
      _Version        =   393216
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
      DataMember      =   "cmdStockSearch"
      ColumnCount     =   5
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "Quantity"
         Caption         =   "Quantity"
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
         Caption         =   "Price"
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
      BeginProperty Column04 
         DataField       =   "ID"
         Caption         =   "ID"
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
            ColumnWidth     =   3825.071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   959.811
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtKeyword 
      Height          =   360
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type in any keyword:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   300
      Width           =   1695
   End
End
Attribute VB_Name = "frmStockSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSearch_Click()
    'Recordset must be closed before executed with parameters
    If dePNC.rscmdStockSearch.State = adStateOpen Then
        dePNC.rscmdStockSearch.Close
    End If
    'Execute cmdStockSearch with required parameters
    dePNC.cmdStockSearch "%" & txtKeyword.Text & "%", "%" & txtKeyword.Text & "%"
    With dgStock
        .DataMember = "cmdStockSearch"
        Set .DataSource = dePNC
        .Refresh
    End With
    'Check if result was returned
    If dePNC.rscmdStockSearch.EOF Then
        MsgBox "Sorry, no match was found.", vbInformation, "No match"
        txtKeyword.SetFocus
        Exit Sub
    End If
End Sub

Private Sub cmdSellThisItem_Click()
    'Check if recordset is open
    If dePNC.rscmdStockSearch.State = adStateClosed Then
        MsgBox "Start by doing a search on the stock", vbInformation, "Stock Search"
        Exit Sub
    End If
    'Check if item has been selected
    If dePNC.rscmdStockSearch.EOF Or dePNC.rscmdStockSearch.BOF Then
        MsgBox "No selected item from the grid has been given.  Select one or do another search.", vbInformation, "No Item Selected"
        Exit Sub
    End If
    'Determine ID
    Dim ID As Integer
    ID = Mid(dePNC.rscmdStockSearch.Fields("ID").Value, 7)
    'Add another row in frmStockSale
    frmStockSale.cmdAddAnotherItem_Click
    If Left(dePNC.rscmdStockSearch.Fields("ID").Value, 1) = "P" Then
        'Show PawnStock Item on frmStockSale
        frmStockSale.cboPawnStockID(frmStockSale.row).Text = ID
    Else 'Show Normal/Standard Stock Item on frmStockSale
        frmStockSale.cboNormalStockID(frmStockSale.row).Text = ID
    End If
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Close recordset if open, to prevent current results from displaying at next Show
    If dePNC.rscmdStockSearch.State = adStateOpen Then
        dePNC.rscmdStockSearch.Close
    End If
    'fraStockSell.Visible = False
End Sub
