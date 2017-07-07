VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUtilities 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlg 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Get path to Access database
Public Function GetDBPath() As String
    With dlg
        ' Set CancelError is True
        .CancelError = True
        .DialogTitle = "Locate the Access Database"
        On Error GoTo ErrHandler
        ' Set flags
        .Flags = cdlOFNHideReadOnly
        ' Set filters
        .Filter = "Access Database Files (*.mdb)|*.mdb"
        ' Specify default filter
        .FilterIndex = 1
        ' Display the Open dialog box
        .ShowOpen
        ' Save selection to registry
        SaveSetting "Cash&BuyBack", "Settings", "DBLocation", .FileName
        ' Return name of selected file
        GetDBPath = .FileName
        Exit Function
    End With
  
ErrHandler:
    Select Case Err.Number
        Case 32755 'User pressed the Cancel button
            GetDBPath = "not set"
            'MsgBox "You cancelled setting the path to the database file." & vbCrLf & "No changes was applied.", vbInformation
    End Select
End Function




