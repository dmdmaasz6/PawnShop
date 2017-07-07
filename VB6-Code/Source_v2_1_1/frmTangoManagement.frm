VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTangoManagement 
   Caption         =   "Tango Stock Mangement"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8100
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTangoManagement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   8100
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Update:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   7575
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTangoManagement.frx":08CA
         Height          =   615
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   7095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Delete: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   240
      TabIndex        =   7
      Top             =   5280
      Width           =   7575
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   550
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "To delete one of the entries, select it and click on Delete."
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   425
         Width           =   4995
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Add: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   7575
      Begin VB.TextBox txtRecommendedSellingPrice 
         DataField       =   "RecommendedSellingPrice"
         DataMember      =   "cmdAllCellphoneAssessories"
         DataSource      =   "dePNC"
         Height          =   360
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         Width           =   1320
      End
      Begin VB.TextBox txtDescription 
         DataField       =   "Description"
         DataMember      =   "cmdAllCellphoneAssessories"
         DataSource      =   "dePNC"
         Height          =   360
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   3975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Default         =   -1  'True
         Height          =   550
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   1495
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Price:"
         Height          =   240
         Index           =   2
         Left            =   165
         TabIndex        =   6
         Top             =   900
         Width           =   1170
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   240
         Index           =   1
         Left            =   270
         TabIndex        =   5
         Top             =   405
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   550
      Left            =   2880
      TabIndex        =   0
      Top             =   6720
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc adodcCellphone 
      Height          =   375
      Left            =   4320
      Top             =   6360
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Caption         =   "Adodc1"
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
   Begin MSDataGridLib.DataGrid dgCellphone 
      Bindings        =   "frmTangoManagement.frx":095B
      Height          =   2415
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   16777215
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "NSID"
         Caption         =   "NSID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
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
         DataField       =   "RecommendedSellingPrice"
         Caption         =   "Selling Price"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "QuantityAvailable"
         Caption         =   "Qty Available"
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
         DataField       =   "QuantityLayBuy"
         Caption         =   "Qty on LayBuy"
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
            Locked          =   -1  'True
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1200.189
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTangoManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    'Ensure information has been specified
    If Trim(txtDescription.Text) = "" Then
        MissingInfo ("Description")
        Exit Sub
    ElseIf Trim(txtRecommendedSellingPrice.Text) = "" Then
        MissingInfo ("Selling Price")
        Exit Sub
    End If
    'Add new grouped stock item
    With adodcCellphone.Recordset
        .AddNew
        .Fields("Description") = txtDescription.Text
        .Fields("RecommendedSellingPrice") = txtRecommendedSellingPrice.Text
        .Fields("CellphoneAssessories") = 2
        .Fields("Quantity") = 0
        .Fields("QuantitySold") = 0
        .Fields("QuantityLayBuy") = 0
        .Update
    End With
    'Clear fields
    txtDescription.Text = ""
    txtRecommendedSellingPrice.Text = ""
End Sub

Private Sub cmdDelete_Click()
    'Confirm deletion
    Dim intResponse As Integer
    intResponse = MsgBox("Do you really want to delete this entry?", vbYesNo, "Confirm Cancellation")
    If intResponse = vbYes Then
        adodcCellphone.Recordset.Delete
        adodcCellphone.Recordset.Update
    End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
    ' Repopulate ADODC Control
    frmTangoManagement.adodcCellphone.ConnectionString = gstrConnect
    frmTangoManagement.adodcCellphone.RecordSource = "SELECT NSID, Description, RecommendedSellingPrice, (Quantity - QuantitySold - QuantityLayBuy) as QuantityAvailable, Quantity, QuantitySold, QuantityLayBuy, CellphoneAssessories FROM NormalStock WHERE (CellphoneAssessories = 2) order by Description"
    frmTangoManagement.adodcCellphone.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Ensure all changes have been updated
    On Error Resume Next
    If adodcCellphone.Recordset.State = adStateOpen Then
        adodcCellphone.Recordset.Update
    End If
End Sub

Private Sub txtRecommendedSellingPrice_KeyPress(KeyAscii As Integer)
    'Ensure that only numbers, backspace or decimal point are entered
    If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 46) Then
        MsgBox "Price must have a numeric value.", vbInformation, "Key invalid"
        KeyAscii = 0
    End If
End Sub

Private Sub txtRecommendedSellingPrice_LostFocus()
    SetCurr txtRecommendedSellingPrice
End Sub

