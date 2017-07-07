VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCustomerHistory 
   Caption         =   "Customer History"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10665
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomerHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10665
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   " History "
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
      Height          =   5415
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   10215
      Begin MSComctlLib.ListView lstHsitory 
         Height          =   5055
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   8916
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PnCID"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Serial Nr"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Month"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Time"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "DateHandedIn"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "BuyBackAmount"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ExpiryDate"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Status"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "ActualBuyBackAmount"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "BuyBackDate"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "ExpiryDate (7WorkDays)"
            Object.Width           =   706
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Customer "
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
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin VB.TextBox txtCustomerName 
         Height          =   360
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   300
         Width           =   3495
      End
      Begin VB.TextBox txtCustID 
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Customer Name:"
         Height          =   240
         Left            =   4035
         TabIndex        =   3
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Customer ID:"
         Height          =   240
         Left            =   435
         TabIndex        =   2
         Top             =   360
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmCustomerHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCustID As String
Public strCustName As String

Private Sub cmdOK_Click()
Unload Me

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim z As Integer
Dim arrItems(0 To 10000) As String
Dim iListItem As Integer

MousePointer = vbHourglass

txtCustID.Text = strCustID
txtCustomerName.Text = strCustName

i = 0

    rs.Source = "select " & _
                "PnCID, " & _
                "MonthNr," & _
                "PTTime," & _
                "PurchaseAmount, " & _
                "DatehandedIn, " & _
                "BuyBAckAmount, " & _
                "ExpiryDate, " & _
                "Status, " & _
                "ActualBuyBackAmount, " & _
                "BuyBackDate, " & _
                "ExpiryDateAfter7WorkDays " & _
                "from PawningTransactions where custID  = " & strCustID

    rs.Open
    Do Until rs.EOF
        If IsNull(rs.Fields("PncID").Value) Then
            lstHsitory.ListItems.Add , , " "
            lstHsitory.ListItems(i).SubItems(1) = "No History Data ..."
        Else
            'cboCashRegister.AddItem UCase(rs.Fields("cashregister").Value)
            'frmCashRegister.cboCashRegister.AddItem UCase(rs.Fields("cashregister").Value)
            lstHsitory.ListItems.Add , , rs.Fields("PnCID").Value
            i = i + 1
            arrItems(i) = rs.Fields("PnCID").Value
            If Not IsNull(rs.Fields("MonthNr").Value) Then lstHsitory.ListItems(i).SubItems(3) = rs.Fields("MonthNr").Value
            If Not IsNull(rs.Fields("PTTime").Value) Then lstHsitory.ListItems(i).SubItems(4) = rs.Fields("PTTime").Value
            If Not IsNull(rs.Fields("PurchaseAmount").Value) Then lstHsitory.ListItems(i).SubItems(5) = rs.Fields("PurchaseAmount").Value
            If Not IsNull(rs.Fields("DatehandedIn").Value) Then lstHsitory.ListItems(i).SubItems(6) = rs.Fields("DatehandedIn").Value
            If Not IsNull(rs.Fields("BuyBAckAmount").Value) Then lstHsitory.ListItems(i).SubItems(7) = rs.Fields("BuyBAckAmount").Value
            If Not IsNull(rs.Fields("ExpiryDate").Value) Then lstHsitory.ListItems(i).SubItems(8) = rs.Fields("ExpiryDate").Value
            If Not IsNull(rs.Fields("Status").Value) Then lstHsitory.ListItems(i).SubItems(9) = rs.Fields("Status").Value
            If Not IsNull(rs.Fields("ActualBuyBackAmount").Value) Then lstHsitory.ListItems(i).SubItems(10) = rs.Fields("ActualBuyBackAmount").Value
            If Not IsNull(rs.Fields("BuyBackDate").Value) Then lstHsitory.ListItems(i).SubItems(11) = rs.Fields("BuyBackDate").Value
            If Not IsNull(rs.Fields("ExpiryDateAfter7WorkDays").Value) Then lstHsitory.ListItems(i).SubItems(12) = rs.Fields("ExpiryDateAfter7WorkDays").Value
        End If
        rs.MoveNext
    Loop
        
    'Ensure that recordset is closed
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    iListItem = 1
    
    If i <> 0 Then
        For z = 1 To i
            rs.Source = "select Description,SerialNr From PawnStock where PnCID = " & arrItems(iListItem)
            rs.Open
            Do Until rs.EOF
                If Not IsNull(rs.Fields("Description").Value) Then lstHsitory.ListItems(iListItem).SubItems(1) = rs.Fields("Description").Value
                If Not IsNull(rs.Fields("SerialNr").Value) Then lstHsitory.ListItems(iListItem).SubItems(2) = rs.Fields("SerialNr").Value
                rs.MoveNext
            Loop
            
            
            If rs.State = adStateOpen Then
                rs.Close
            End If
    
            iListItem = iListItem + 1
            
        Next z
    Else
        lstHsitory.ListItems.Add , , " "
        lstHsitory.ListItems(1).SubItems(1) = "No History Data ..."
    End If
MousePointer = vbNormal
    
LV_AutoSizeColumn lstHsitory

End Sub

Public Sub LV_AutoSizeColumn(LV As ListView, _
                Optional Column As ColumnHeader = Nothing)
Dim col2adjust As Long

For col2adjust = 0 To LV.ColumnHeaders.Count - 1

   Call SendMessage(LV.hWnd, _
                    LVM_SETCOLUMNWIDTH, _
                    col2adjust, _
                    ByVal LVSCW_AUTOSIZE_USEHEADER)
Next

LV.Refresh
 

End Sub

