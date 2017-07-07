VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOtherExpenses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Other Expenses"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOtherExpenses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4770
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   550
      Left            =   2640
      TabIndex        =   9
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdProcessExpense 
      Caption         =   "Process"
      Height          =   550
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   1680
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtDescription 
      Height          =   975
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   5
      Top             =   1185
      Width           =   2895
   End
   Begin VB.ComboBox cboCategory 
      Height          =   360
      ItemData        =   "frmOtherExpenses.frx":08CA
      Left            =   1680
      List            =   "frmOtherExpenses.frx":08F5
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   630
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   635
      _Version        =   393216
      Format          =   47644672
      CurrentDate     =   37066
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
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
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
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
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   200
      Width           =   975
   End
End
Attribute VB_Name = "frmOtherExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdprocessExpense_Click()
    ' Input validation
    If cboCategory.Text = "" Then
        MsgBox "Please specify a category.", vbInformation, "Missing Information"
        Exit Sub
    ElseIf Trim(txtAmount.Text) = "" Then
        MsgBox "Please specify an amount.", vbInformation, "Missing Information"
        Exit Sub
    End If
    
    ' Prepare Values
    Dim strDescription As String
    Dim intIndex As Integer
    If Trim(txtDescription.Text) = "" Then 'Error generated in insert statement when txtDescription is empty
        strDescription = "No description"
    Else
        strDescription = txtDescription.Text
    End If
    intIndex = cboCategory.ListIndex + 1 'ListIndex starts at 0
    
    'dePNC.cmdOtherExpenses dtpDate.Value, intIndex, strDescription, -(CCur(txtAmount.Text))
    dePNC.cmdInsertCashTransaction gstrCR, "Categorized Expenses - index at top of report", dtpDate.Value, 0, 0, 0, 0, intIndex, strDescription, -(CCur(txtAmount.Text)), 0, 0, 0
    Unload Me
End Sub

Private Sub Form_Load()
    dtpDate.Value = Date
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
