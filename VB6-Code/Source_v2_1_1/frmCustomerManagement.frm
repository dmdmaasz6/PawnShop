VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCustomerManagement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Management"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomerManagement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   10200
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdUseThisCustomer 
      Caption         =   "Use this Customer"
      Height          =   550
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdShowHistory 
      Caption         =   "Show Customer History ..."
      Height          =   550
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H80000004&
      Caption         =   "E&xit"
      Height          =   550
      Left            =   6480
      TabIndex        =   11
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   " Details "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   9975
      Begin VB.CommandButton cmdSearchCustomer 
         Caption         =   "Search"
         Height          =   550
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2040
         Width           =   2295
      End
      Begin VB.CommandButton cmdUpdateCustomerInformation 
         Caption         =   "Update Information"
         Height          =   550
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CommandButton cmdAddCustomer 
         Caption         =   "Add New"
         Height          =   550
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton cmdClearFields 
         Caption         =   "Clear Fields"
         Height          =   550
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtAddress 
         Height          =   2235
         Left            =   4440
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtTelW 
         Height          =   360
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtTelH 
         Height          =   360
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   3
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtName 
         Height          =   360
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   0
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtIDnr 
         Height          =   360
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   1
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtCustID 
         Height          =   360
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   15
         ToolTipText     =   "Cash&BuyBack ID"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Address:"
         Height          =   255
         Left            =   3120
         TabIndex        =   26
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel(w):"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel(h):"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2325
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ID number:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1365
         Width           =   1215
      End
      Begin VB.Label lblCustID 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Cash&BuyBack ID"
         Top             =   400
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid dgCustomerList 
      Bindings        =   "frmCustomerManagement.frx":08CA
      Height          =   3615
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   6376
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
      DataMember      =   "cmdSearchCustomer"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "CustID"
         Caption         =   "CBB ID"
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
         DataField       =   "IDNr"
         Caption         =   "IDNr"
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
      BeginProperty Column03 
         DataField       =   "Address"
         Caption         =   "Address"
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
         DataField       =   "TelH"
         Caption         =   "TelH"
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
         DataField       =   "TelW"
         Caption         =   "TelW"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCustomerUpdate 
      BackColor       =   &H80000004&
      Caption         =   "Customer Update: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6720
      TabIndex        =   17
      Top             =   4920
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "To update customer information, make the necessary changes and then click on Update Customer Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000004&
      Caption         =   "Customer Search: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3480
      TabIndex        =   14
      Top             =   120
      Width           =   4095
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Click on any one of the customers below to have his/hers information display to the left."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmCustomerManagement.frx":08DE
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame fraNewCustomer 
      BackColor       =   &H80000004&
      Caption         =   "New Customer: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3600
      TabIndex        =   12
      Top             =   5040
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "To add a new customer, complete the fields and click on Add New Customer button."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmCustomerManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAddCustomer_Click()
    'validate input
    If Trim(txtName.Text) = "" Then
        MsgBox "You must specify a name.", vbInformation, "Customer Information"
        Exit Sub
    Else
        'add customer
        'check for duplicate customers
        rs.Source = "Select * from Customers where Name like '%" & Trim(txtName.Text) & "%' or IDnr like '%" & Trim(txtIDNr.Text) & "%'"
        rs.Open
        If Not rs.EOF Then
            'duplicate entry
            Dim intResponse As Integer
            intResponse = MsgBox("A customer with similar details already exists." & vbCrLf & vbCrLf & _
                "Name: " & rs.Fields("Name") & vbCrLf & _
                "Customer ID: " & rs.Fields("CustID") & vbCrLf & _
                "RSA ID: " & rs.Fields("IDnr") & vbCrLf & _
                "Tel(H): " & rs.Fields("TelH") & vbCrLf & _
                "Tel(W): " & rs.Fields("TelW") & vbCrLf & _
                "Address: " & rs.Fields("Address") & vbCrLf & vbCrLf & "Do you still wish to add a new customer?" _
                , vbInformation + vbYesNo + vbDefaultButton2, "Duplicate Names")
            rs.Close
            If intResponse = vbNo Then
                'cancel new customer
                EmptyTextboxes
                Exit Sub
            End If
        End If
        'Ensure that recordset is closed
        If rs.State = adStateOpen Then
            rs.Close
        End If
        'add new customer
        dePNC.cmdNewCustomer Trim(txtName.Text), Trim(txtIDNr.Text), Trim(txtTelW.Text), Trim(txtTelH.Text), Trim(txtAddress.Text)
        'display customerid of new customer
        With dePNC.rscmdLastCustID
            .Open
            Dim custID As Integer
            custID = .Fields("LastCustID").Value
            MsgBox "CustomerID of new customer is: " & custID, vbInformation, "New Customer ID"
            .Close
        End With
        EmptyTextboxes
    End If
End Sub

Private Sub cmdClearFields_Click()
    'Clear text fields
    txtAddress.Text = ""
    txtCustID.Text = ""
    txtIDNr.Text = ""
    txtName.Text = ""
    txtTelH.Text = ""
    txtTelW.Text = ""
    'Clear dgCustomerList
    'Close recordset, if open, so that recordset is not taken over to successive display
    If dePNC.rscmdSearchCustomer.State = adStateOpen Then
        dePNC.rscmdSearchCustomer.Close
    End If
    With dgCustomerList
        .DataMember = "cmdSearchCustomer"
        Set .DataSource = dePNC
        .Refresh
    End With
    txtName.SetFocus
End Sub

Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub cmdSearchCustomer_Click()
    ' You must close the recordset before changing the parameter.
    If dePNC.rscmdSearchCustomer.State = adStateOpen Then
        dePNC.rscmdSearchCustomer.Close
    End If
    'If txtCustID is empty, cmdSearchCustomer generates an error
    If txtCustID.Text = "" Then txtCustID.Text = 0
    'Execute cmdSearchCustomer with all parameters, ie reopen recordset
    dePNC.cmdSearchCustomer txtCustID.Text, "%" & txtName.Text & "%", _
        "%" & txtIDNr.Text & "%", "%" & txtTelW.Text & "%", _
        "%" & txtTelH.Text & "%", "%" & txtAddress.Text & "%"
    'Check if any result as returned
    If dePNC.rscmdSearchCustomer.EOF Then
        MsgBox "Sorry, no match was found.", vbInformation, "No match"
        cmdClearFields_Click
        txtName.SetFocus
        Exit Sub
    End If
    With dgCustomerList
        .DataMember = "cmdSearchCustomer"
        Set .DataSource = dePNC
    End With
    dgCustomerList_Click
End Sub

Private Sub cmdShowHistory_Click()
If txtCustID.Text <> "" Then
    frmCustomerHistory.strCustID = txtCustID.Text
    frmCustomerHistory.strCustName = txtName.Text
    Load frmCustomerHistory
    frmCustomerHistory.Show vbModal
Else
    MsgBox "Please select Customer!", vbInformation + vbOKOnly, App.EXEName
End If

End Sub

Private Sub cmdUpdateCustomerInformation_Click()
    'Ensure that record is displayed
    If Trim(txtName.Text) = "" Then
        cmdSearchCustomer_Click
        dgCustomerList_Click
    End If
    
    'Seek confirmation
    Dim intResponse As Integer
    intResponse = MsgBox("Do you really want to update the customer information?", vbYesNo + vbDefaultButton2 + vbInformation, "Confirm update")
    If intResponse = vbNo Then Exit Sub
    
    ' update recordset
    With dePNC.rscmdSearchCustomer
        .Fields("name").Value = txtName.Text
        .Fields("idnr").Value = txtIDNr.Text
        .Fields("telw").Value = txtTelW.Text
        .Fields("telh").Value = txtTelH.Text
        .Fields("address").Value = txtAddress.Text
        ' update underlying database
        .Update
    End With
End Sub
' To take selected customer information to frmPawningOut
Private Sub cmdUseThisCustomer_Click()
    'Ensure that a customer has been selected
    If Trim(txtCustID.Text) = "" Then
        MissingInfo ("Customer")
        Exit Sub
    End If
    frmPawningOut.txtCustomerID.Text = txtCustID.Text
    frmPawningOut.cmdGetDetails_Click
    Unload Me
End Sub

Private Sub dgCustomerList_Click()
    ' Ensure that recordset has been opened
    If dePNC.rscmdSearchCustomer.State = adStateClosed Then
        Exit Sub
    End If
    With dePNC.rscmdSearchCustomer
        ' Check if recordset has records
        If .EOF Then Exit Sub
        ' display selected customer information in textboxes
        txtCustID.Text = .Fields("custid").Value
        txtName.Text = .Fields("name").Value
        txtIDNr.Text = .Fields("idnr").Value
        txtTelW.Text = .Fields("telw").Value
        txtTelH.Text = .Fields("telh").Value
        txtAddress.Text = .Fields("address").Value
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Close recordset, if open, so that recordset is not taken over to successive display
    If dePNC.rscmdSearchCustomer.State = adStateOpen Then
        dePNC.rscmdSearchCustomer.Close
    End If
End Sub

Private Sub EmptyTextboxes()
    ' Empty textboxes
    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        End If
    Next
End Sub

