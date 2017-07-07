VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStockPurchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Purchase"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStockPurchase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   6285
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optTango 
      Caption         =   "Tango Stock"
      Height          =   315
      Left            =   4440
      TabIndex        =   30
      Top             =   150
      Width           =   1695
   End
   Begin VB.OptionButton optNormal 
      Caption         =   "Normal Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   26
      Top             =   150
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton optStand 
      Caption         =   "Standard Stock"
      Height          =   315
      Left            =   2160
      TabIndex        =   25
      Top             =   150
      Width           =   2175
   End
   Begin VB.CommandButton cmdPurchase 
      Caption         =   "Purchase"
      Height          =   550
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   550
      Left            =   3360
      TabIndex        =   23
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CheckBox chkStandardStock 
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
      Left            =   6000
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Caption         =   " Transaction: "
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
      Height          =   2295
      Left            =   480
      TabIndex        =   10
      Top             =   4320
      Width           =   5415
      Begin VB.TextBox txtQuantity 
         Height          =   360
         Left            =   2160
         TabIndex        =   4
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtRecommendedSellingPrice 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2160
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtPurchasePrice 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2160
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   360
         Left            =   2160
         TabIndex        =   5
         Top             =   840
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         Format          =   47579136
         CurrentDate     =   37066
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Selling Price:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1860
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Purchase Price:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1395
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Date:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   915
         Width           =   615
      End
   End
   Begin VB.Frame fraTango 
      Caption         =   " Tango Stock: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3495
      Left            =   480
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   5415
      Begin MSDataGridLib.DataGrid dgTangoStock 
         Bindings        =   "frmStockPurchase.frx":08CA
         Height          =   2775
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         DataMember      =   "cmdTangoStockCodeList"
         ColumnCount     =   4
         BeginProperty Column00 
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
            Caption         =   "Price"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3195.213
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   14.74
            EndProperty
         EndProperty
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Click in the margin next to the item purchased"
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
         Left            =   360
         TabIndex        =   29
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame fraStandardStock 
      Caption         =   " Standard Stock: "
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
      Height          =   3495
      Left            =   480
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   5415
      Begin MSDataGridLib.DataGrid dgStandardStock 
         Bindings        =   "frmStockPurchase.frx":08DE
         Height          =   2775
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         DataMember      =   "cmdStandardStockCodeList"
         ColumnCount     =   4
         BeginProperty Column00 
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
            Caption         =   "Price"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3195.213
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   14.74
            EndProperty
         EndProperty
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Click in the margin next to the item purchased"
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
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame fraNormalStock 
      Caption         =   " Normal Stock: "
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
      Height          =   3495
      Left            =   480
      TabIndex        =   8
      Top             =   600
      Width           =   5415
      Begin VB.TextBox txtDescription 
         Height          =   975
         Left            =   1560
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtAddress 
         Height          =   975
         Left            =   1560
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtIDNr 
         Height          =   360
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   1
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtName 
         Height          =   360
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Description:"
         Height          =   615
         Left            =   240
         TabIndex        =   19
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Address:"
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Stock"
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
      Left            =   6360
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmStockPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkStandardStock_Click()
    'Display correct frames
    ' Standard Stock
    If chkStandardStock Then
        fraStandardStock.Visible = True
        fraNormalStock.Visible = False
        ' Populate Recommended Selling Price
        If Not dePNC.rscmdStandardStockCodeList.EOF Then
            dgStandardStock_Click
        End If
    Else 'Normal Stock
        fraStandardStock.Visible = False
        fraNormalStock.Visible = True
        ' Empty Recommended Selling Price
        txtRecommendedSellingPrice.Text = ""
    End If
    txtPurchasePrice.Text = ""
    txtQuantity.Text = 1
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPurchase_Click()
    ' Input Validation
    ' Standard Stock
    'If chkStandardStock.Value Then  'old code
    If optStand.Value = True Then
        ' No validation
    ElseIf optTango.Value = True Then
        ' No validation
    Else    'Normal Stock
        If Trim(txtName.Text) = "" Then
            MissingInfo "a name for the supplier"
            txtName.SetFocus
            Exit Sub
        ElseIf Trim(txtDescription.Text) = "" Then
            MissingInfo "a description for the stock item"
            txtDescription.SetFocus
            Exit Sub
        End If
    End If
    ' Normal and Standard Stock
    If Trim(txtPurchasePrice.Text) = "" Then
        MissingInfo "a purchase price for the stock item"
        txtPurchasePrice.SetFocus
        Exit Sub
    ElseIf Trim(txtRecommendedSellingPrice.Text) = "" Then
        MissingInfo "a recommended selling price for the stock item"
        txtRecommendedSellingPrice.SetFocus
        Exit Sub
    End If
    
    ' Insert stock purchase into NormalStock table
    'If chkStandardStock.Value Then 'old code
    If optStand.Value = True Then
        'Update Quantity and RecommendedSellingPrice
        cmd.CommandText = "Update NormalStock set RecommendedSellingPrice = " & txtRecommendedSellingPrice.Text & ", Quantity = Quantity + " & txtQuantity.Text & " where NSID = " & dePNC.rscmdStandardStockCodeList.Fields("NSID").Value
        cmd.Execute
    ElseIf optTango.Value = True Then
        'Update Quantity and RecommendedSellingPrice
        cmd.CommandText = "Update NormalStock set RecommendedSellingPrice = " & txtRecommendedSellingPrice.Text & ", Quantity = Quantity + " & txtQuantity.Text & " where NSID = " & dePNC.rscmdTangoStockCodeList.Fields("NSID").Value
        cmd.Execute
    Else 'Normal Stock
        dePNC.cmdStockPurchaseToNormalStock CDate(dtpDate.Value), txtName.Text, txtIDNr.Text, txtAddress.Text, txtDescription.Text, CInt(txtQuantity.Text), CCur(txtPurchasePrice.Text), CCur(txtRecommendedSellingPrice.Text), 0
    End If
    
    ' Retrieve NSID from lastly added item and display
    'Normal Stock
    'If chkStandardStock.Value = 0 Then 'old code
    If optNormal.Value = True Then
        Dim rs As ADODB.Recordset
        Dim CurrNSID As Integer
        Set rs = New ADODB.Recordset
        rs.Open "Select max(NSID) as MaxNSID from NormalStock", dePNC.cnPNC
        CurrNSID = rs.Fields("MaxNSID").Value
        rs.Close
        MsgBox "Mark this stock item(s) with the following Normal Stock ID (NSID): " & CurrNSID, vbInformation, "Mark Stock Items"
    End If
    
    ' Insert Stock Purchase into CashTransactions table
    'If chkStandardStock.Value Then 'old code
    If optStand.Value = True Then
        ' Standard Stock Purchase
        dePNC.cmdInsertCashTransaction gstrCR, "Standard Stock Purchase", dtpDate.Value, 0, 0, txtQuantity.Text, dePNC.rscmdStandardStockCodeList.Fields("NSID").Value, 0, dePNC.rscmdStandardStockCodeList.Fields("Description").Value, -((CCur(txtPurchasePrice.Text)) * txtQuantity.Text), txtPurchasePrice.Text, 0, 0
    ElseIf optTango.Value = True Then
        ' Standard Stock Purchase
        dePNC.cmdInsertCashTransaction gstrCR, "Tango Stock Purchase", dtpDate.Value, 0, 0, txtQuantity.Text, dePNC.rscmdTangoStockCodeList.Fields("NSID").Value, 0, dePNC.rscmdTangoStockCodeList.Fields("Description").Value, -((CCur(txtPurchasePrice.Text)) * txtQuantity.Text), txtPurchasePrice.Text, 0, 0
    Else
        ' Normal Stock Purchase
        dePNC.cmdInsertCashTransaction gstrCR, "Normal Stock Purchase", dtpDate.Value, 0, 0, txtQuantity.Text, CurrNSID, 0, txtDescription.Text, -((CCur(txtPurchasePrice.Text)) * txtQuantity.Text), txtPurchasePrice.Text, 0, 0
    End If
    Unload Me
End Sub

Private Sub dgStandardStock_Click()
    On Error Resume Next
    'Populate Recommended Selling Price
    txtRecommendedSellingPrice.Text = dePNC.rscmdStandardStockCodeList.Fields("RecommendedSellingPrice")
    txtRecommendedSellingPrice_LostFocus
End Sub

Private Sub dgTangoStock_Click()
    On Error Resume Next
    'Populate Recommended Selling Price
    txtRecommendedSellingPrice.Text = dePNC.rscmdTangoStockCodeList.Fields("RecommendedSellingPrice")
    txtRecommendedSellingPrice_LostFocus
    
End Sub

Private Sub Form_Load()
    dtpDate.Value = Date
End Sub

Private Sub optNormal_Click()
If optNormal.Value = True Then
    optNormal.FontBold = True
    optStand.FontBold = False
    optTango.FontBold = False
    fraStandardStock.Visible = False
    fraNormalStock.Visible = True
    fraTango.Visible = False
End If

End Sub

Private Sub optStand_Click()
If optStand.Value = True Then
    optNormal.FontBold = False
    optStand.FontBold = True
    optTango.FontBold = False
    fraStandardStock.Visible = True
    fraNormalStock.Visible = False
    fraTango.Visible = False
End If

End Sub

Private Sub optTango_Click()
If optTango.Value = True Then
    optNormal.FontBold = False
    optStand.FontBold = False
    optTango.FontBold = True
    fraStandardStock.Visible = False
    fraNormalStock.Visible = False
    fraTango.Visible = True
End If

End Sub

Private Sub txtPurchasePrice_KeyPress(KeyAscii As Integer)
    'Ensure that only numbers, backspace or decimal point are entered
    If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 46) Then
        MsgBox "Purchase Price must have a numeric value.", vbInformation, "Key invalid"
        KeyAscii = 0
    End If
End Sub

Private Sub txtPurchasePrice_LostFocus()
    SetCurr txtPurchasePrice
End Sub

Private Sub txtPurchasePrice_Validate(Cancel As Boolean)
    'Check if value is valid
    If Not IsNumeric(txtPurchasePrice.Text) Then
        MsgBox "The Purchase Price must have a numeric value.", vbInformation, "Number required"
        Cancel = True
    Else
        If CCur(txtPurchasePrice.Text) < 0 Then
            MsgBox "The Purchase Price must have a value greater than 0.", vbInformation, "Number required"
            Cancel = True
        End If
    End If
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
    'Ensure that only numbers, backspace or decimal point are entered
    If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 46) Then
        MsgBox "Quantity must have a numeric value.", vbInformation, "Key invalid"
        KeyAscii = 0
    End If
End Sub

Private Sub txtQuantity_Validate(Cancel As Boolean)
    'Check if value is valid
    If Not (IsNumeric(txtQuantity.Text) And CCur(txtQuantity.Text) > 0) Then
        MsgBox "The Quantity must have a positive numeric value.", vbInformation, "Number required"
        Cancel = True
    End If
End Sub

Private Sub txtRecommendedSellingPrice_KeyPress(KeyAscii As Integer)
    'Ensure that only numbers, backspace or decimal point are entered
    If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 46) Then
        MsgBox "Recommended Selling Price must have a numeric value.", vbInformation, "Key invalid"
        KeyAscii = 0
    End If
End Sub

Private Sub txtRecommendedSellingPrice_LostFocus()
    SetCurr txtRecommendedSellingPrice
End Sub

Private Sub txtRecommendedSellingPrice_Validate(Cancel As Boolean)
    'Check if value is valid
    If Not IsNumeric(txtRecommendedSellingPrice.Text) Then
        MsgBox "The Recommended Selling Price must have a numeric value.", vbInformation, "Number required"
        Cancel = True
    Else
        If CCur(txtRecommendedSellingPrice.Text) < 0 Then
            MsgBox "The Recommended Selling Price must have a value greater than 0.", vbInformation, "Number required"
            Cancel = True
        End If
    End If
End Sub

