VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLayBuySearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search for Lay Byes"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   Icon            =   "frmLayBuySearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7725
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLayByeTransaction 
      Caption         =   "Lay Bye Transaction"
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   4080
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H80000004&
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   495
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame fraStockSell 
      Caption         =   "To Select a Lay Bye Transaction:"
      Height          =   975
      Left            =   -4080
      TabIndex        =   1
      Top             =   4680
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Transaction by clicking in the margin next to it and click on Lay Bye Transaction."
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4455
      End
   End
   Begin MSDataGridLib.DataGrid dgLayByes 
      Bindings        =   "frmLayBuySearch.frx":08CA
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   14737632
      HeadLines       =   1
      RowHeight       =   15
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "cmdLayByeSearch"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "LBID"
         Caption         =   "LB ID"
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
         DataField       =   "Name"
         Caption         =   "Customer Name"
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
         DataField       =   "Description"
         Caption         =   "Stock Description"
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
         DataField       =   "NSID"
         Caption         =   "NS ID"
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
         DataField       =   "PSID"
         Caption         =   "PS ID"
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
            Alignment       =   1
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3179.906
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   734.74
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type in part of the name of the customer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   175
      Width           =   2055
   End
End
Attribute VB_Name = "frmLayBuySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSearch_Click()
'    'Recordset must be closed before executed with parameters
'    If dePNC.rscmdLayByeSearch.State = adStateOpen Then
'        dePNC.rscmdLayByeSearch.Close
'    End If
'    'Execute cmdLayByeSearch with required parameters
'    dePNC.cmdLayByeSearch "%" & txtName.Text & "%"
'    With dgLayByes
'        .DataMember = "cmdLayByeSearch"
'        Set .DataSource = dePNC
'        .Refresh
'    End With
'    'Check if result was returned
'    If dePNC.rscmdLayByeSearch.EOF Then
'        MsgBox "Sorry, no match was found.", vbInformation, "No match"
'        txtName.SetFocus
'        Exit Sub
'    End If
End Sub

Private Sub cmdLayByeTransaction_Click()
'    'Check if recordset is open
'    If dePNC.rscmdLayByeSearch.State = adStateClosed Then
'        MsgBox "Start by doing a search", vbInformation, "Lay Bye Search"
'        Exit Sub
'    End If
'    'Check if item has been selected
'    If dePNC.rscmdLayByeSearch.EOF Or dePNC.rscmdLayByeSearch.BOF Then
'        MsgBox "No selected transaction from the grid has been given.  Select one or do another search.", vbInformation, "No Transaction Selected"
'        Exit Sub
'    End If
'    'Apply select transaction to frmLayBuyPayment
'    frmLayBuyPayment.txtID.Text = dePNC.rscmdLayByeSearch.Fields("LBID").Value
'    frmLayBuyPayment.cmdGetDetails_Click
'    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    'Close recordset if open, to prevent current results from displaying at next Show
'    If dePNC.rscmdLayByeSearch.State = adStateOpen Then
'        dePNC.rscmdLayByeSearch.Close
'    End If
'    'fraStockSell.Visible = False
End Sub

