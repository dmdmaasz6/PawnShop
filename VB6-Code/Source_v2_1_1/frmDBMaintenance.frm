VERSION 5.00
Begin VB.Form frmDBMaintenance 
   Caption         =   "Database Maintenance"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDBMaintenance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5760
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   " Clear Database History: "
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
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5535
      Begin VB.Frame Frame2 
         Caption         =   " Cash Transaction History "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   5295
         Begin VB.CommandButton cmdClearCash 
            Caption         =   "Clear ..."
            Height          =   550
            Left            =   3600
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
         Begin VB.ComboBox cmbCashMonths 
            Height          =   360
            ItemData        =   "frmDBMaintenance.frx":08CA
            Left            =   1080
            List            =   "frmDBMaintenance.frx":08CC
            TabIndex        =   8
            Top             =   325
            Width           =   615
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Keep last "
            Height          =   240
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "months' history ..."
            Height          =   240
            Left            =   1800
            TabIndex        =   10
            Top             =   360
            Width           =   1500
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Pawning History "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Visible         =   0   'False
         Width           =   5295
         Begin VB.ComboBox cmbPawnMonths 
            Height          =   360
            ItemData        =   "frmDBMaintenance.frx":08CE
            Left            =   1080
            List            =   "frmDBMaintenance.frx":08D0
            TabIndex        =   6
            Top             =   325
            Width           =   615
         End
         Begin VB.CommandButton cmdClearPawn 
            Caption         =   "Clear ..."
            Height          =   550
            Left            =   3600
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "months' history ..."
            Height          =   240
            Left            =   1800
            TabIndex        =   5
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Keep last "
            Height          =   240
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   885
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   550
      Left            =   2040
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
End
Attribute VB_Name = "frmDBMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdClearCash_Click()
Dim dMonths As Double
Dim dDate As Date
Dim sCount As String
Dim sMonth As String
Dim sYear As String

On Error GoTo lblErr
MousePointer = vbHourglass

If Trim(cmbCashMonths.Text) = "" Then
    MsgBox "Select a month value", vbInformation + vbOKOnly, "Clear Database History"
    Exit Sub
End If

dMonths = CDbl(cmbCashMonths.Text) * -1
dDate = DateAdd("m", dMonths, Date)

sYear = CStr(Year(dDate))
sMonth = CStr(Month(dDate))

rs.Source = "SELECT count(*) " & _
            "From CashTransactions " & _
            "Where Quantity = 0 " & _
            "and Year(TransactionDate)<" & sYear & _
            "and month(TransactionDate)<" & sMonth
rs.Open

sCount = rs.Fields(0).Value

If rs.State = adStateOpen Then
    rs.Close
End If


If sCount <> "0" Then
    If MsgBox(sCount & " records will be deleted from the database," & vbCrLf & _
             "with a Transaction Date before " & CStr(dDate) & vbCrLf & vbCrLf & _
             "Do you want to continue?", vbQuestion + vbYesNo, "Clear CashTransaction Database") = vbYes Then
        rs.Source = "DELETE * " & _
                    "From CashTransactions " & _
                    "Where Quantity = 0 " & _
                    "and month(TransactionDate)<" & sMonth & _
                    "and year(TransactionDate)<" & sYear
        rs.Open
        
        If rs.State = adStateOpen Then
            rs.Close
        End If
        
        MsgBox sCount & " records deleted from the database", vbExclamation + vbOKOnly, "Clear CashTransaction Database"
        
    End If
Else
    MsgBox "There are " & sCount & " records to be deleted from the database", vbExclamation + vbOKOnly, "Clear CashTransaction Database"
End If

MousePointer = vbNormal
Exit Sub

lblErr:
MsgBox "Error Occured during update of database" & vbCrLf & "Error: " & Err.Description
MousePointer = vbNormal

End Sub

Private Sub cmdClearPawn_Click()
Dim dMonths As Double
Dim dDate As Date
Dim sCount As String

On Error GoTo lblErr
MousePointer = vbHourglass

If Trim(cmbPawnMonths.Text) = "" Then
    MsgBox "Select a month value", vbInformation + vbOKOnly, "Clear Database History"
    Exit Sub
End If

dMonths = CDbl(cmbPawnMonths.Text) * -1
dDate = DateAdd("m", dMonths, Date)

rs.Source = "SELECT count(*) " & _
            "From PawningTransactions " & _
            "WHERE ExpiryDateAfter7WorkDays > " & dDate
rs.Open

sCount = rs.Fields(0).Value

If rs.State = adStateOpen Then
    rs.Close
End If


If sCount <> "0" Then
    If MsgBox(sCount & " records will be deleted from the database," & vbCrLf & _
             "with a Expiry Date After 7 WorkDays before " & CStr(dDate) & vbCrLf & vbCrLf & _
             "Do you want to continue?", vbQuestion + vbYesNo, "Clear PawningTransactions Database") = vbYes Then
        rs.Source = "DELETE " & _
                    "From PawningTransactions " & _
                    "WHERE ExpiryDateAfter7WorkDays > " & dDate
        rs.Open
        
        If rs.State = adStateOpen Then
            rs.Close
        End If
        
        MsgBox sCount & " records deleted from the database", vbExclamation + vbOKOnly, "Clear PawningTransactions Database"
        
    End If
Else
    MsgBox "There are " & sCount & " records to be deleted from the database", vbExclamation + vbOKOnly, "Clear PawningTransactions Database"
End If

MousePointer = vbNormal
Exit Sub

lblErr:
MsgBox "Error Occured during update of database" & vbCrLf & "Error: " & Err.Description
MousePointer = vbNormal

End Sub

Private Sub Form_Load()

cmbPawnMonths.AddItem "3"
cmbPawnMonths.AddItem "6"
cmbPawnMonths.AddItem "9"
cmbPawnMonths.AddItem "12"
cmbPawnMonths.AddItem "18"
cmbPawnMonths.AddItem "24"

cmbCashMonths.AddItem "3"
cmbCashMonths.AddItem "6"
cmbCashMonths.AddItem "9"
cmbCashMonths.AddItem "12"
cmbCashMonths.AddItem "18"
cmbCashMonths.AddItem "24"

End Sub
