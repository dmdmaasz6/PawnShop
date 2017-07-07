Attribute VB_Name = "Utilities"
' Connectionstring for database
Public gstrConnect As String
Public DBPath As String
Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Public cmd As ADODB.Command
Public fso As Scripting.FileSystemObject, ErrorReport
Public gstrCR As String 'stores Cash Register Location from registry
Public gstrCurrency As String 'stores Currency from registry
Public gstrVat As String    'stores VAT NUMBER from registry
Public gstrAddress As String  'stores address from registry
Public gstrAddress1 As String
Public gstrAddress2 As String
Public gstrAddress3 As String

Public sTemplatePath As String
Public sTemplateShop As String


Public Sub Main()
    On Error GoTo ErrorHandle
'    'Check for licensing file
'    Dim fso As Scripting.FileSystemObject
'    Set fso = New Scripting.FileSystemObject
'    If Not (fso.FileExists("c:\opsys.dat")) Then
'        'Unlicensed
'        MsgBox "Sorry, but this is an unlicensed software copy." & vbCrLf & "Contact Eben de Wit for details.", vbCritical, "Unlicensed"
'        End
'    End If
    
    'Create/Open file for ErrorReporting
    Set fso = New Scripting.FileSystemObject
    Set ErrorReport = fso.OpenTextFile(App.Path & "\ErrorReport.txt", ForAppending, True)
    
    'Dymanically create DBPath for shared access
    DBPath = GetSetting("Cash&BuyBack", "Settings", "DBLocation", "not set")
    If DBPath = "not set" Then 'path has not yet been set
        MsgBox "The path to the database has not yet been set." & vbCrLf & "Please make the selection in the next window.", vbInformation, "Database Path"
SetDBPath:
        DBPath = frmUtilities.GetDBPath
        'The setting wasn't applied.  Application terminate
        If DBPath = "not set" Then
            MsgBox "The database path has not been set correctly." & vbCrLf & "The application will now terminate.", vbCritical, "Setting unsuccessfull"
            End
        End If
    End If
    
    gstrConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath & ";Persist Security Info=False"
    dePNC.cnPNC.ConnectionString = gstrConnect
    
    'Set ConnectionString and RecordSource properties of ADODCs on frmExpiredPawnTransactions
    With frmExpiredPawnTransactions.adoExpiredPawnTransactions
        .ConnectionString = gstrConnect
        .RecordSource = "SELECT Customers.Name, PawningTransactions.DateHandedIn, PawningTransactions.ExpiryDateAfter7WorkDays, PawningTransactions.PnCID, PawningTransactions.Status FROM Customers INNER JOIN PawningTransactions ON Customers.CustID = PawningTransactions.CustID WHERE (PawningTransactions.ExpiryDateAfter7WorkDays <= #1/1/1900#) AND (PawningTransactions.Status = 1) ORDER BY Customers.Name, PawningTransactions.DateHandedIn"
        .Refresh
    End With
    With frmExpiredPawnTransactions.adoExpiredPawnStock
        .ConnectionString = gstrConnect
        .RecordSource = "SELECT Quantity, Description, SerialNr, Price, PnCID, Status FROM PawnStock where PnCID = 0"
        .Refresh
    End With
    With frmCellphoneManagement.adodcCellphone
        .ConnectionString = gstrConnect
        .RecordSource = "SELECT NSID, Description, RecommendedSellingPrice, (Quantity - QuantitySold - QuantityLayBuy) as QuantityAvailable, Quantity, QuantitySold, QuantityLayBuy, CellphoneAssessories FROM NormalStock WHERE (CellphoneAssessories = 1) order by Description"
        .Refresh
    End With
    With frmTangoManagement.adodcCellphone
        .ConnectionString = gstrConnect
        .RecordSource = "SELECT NSID, Description, RecommendedSellingPrice, (Quantity - QuantitySold - QuantityLayBuy) as QuantityAvailable, Quantity, QuantitySold, QuantityLayBuy, CellphoneAssessories FROM NormalStock WHERE (CellphoneAssessories = 2) order by Description"
        .Refresh
    End With
    
    'Create and open connection
    Set cn = New ADODB.Connection
    cn.ConnectionString = gstrConnect
    cn.Open
    'Create global recordset
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
    'Create global command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    
    'Get Cash Register Location from registry
    gstrCR = GetSetting("Cash&BuyBack", "Settings", "CRLocation", "not set")
    'Cash Register Location not in registry
    If gstrCR = "not set" Then
        Load frmMain
        MsgBox "The cash register location has not yet been set." & vbCrLf & "Please make the selection in the next window.", vbInformation, "Cash Register Location"
        gstrCR = frmMain.SetCashRegister
        If gstrCR = "not set" Then
            MsgBox "The cash register location has not been changed.", vbInformation, "Cash Register Location"
        End If
    End If
    
''    'Get Currency from registry
''    gstrCurrency = GetSetting("Cash&BuyBack", "Settings", "Currency", "not set")
''    If gstrCurrency = "not set" Then
''        Load frmMain
''        MsgBox "The currency symbol has not yet been set." & vbCrLf & "Please make the selection in the next window.", vbInformation, "Currency"
''        gstrCurrency = frmMain.SetCurrency
''        If gstrCurrency = "not set" Then
''            MsgBox "The currency symbol has not been changed.", vbInformation, "Currency"
''        End If
''    End If
''
''    'Get VAT Number from registry
''    gstrVat = GetSetting("Cash&BuyBack", "Settings", "VatNumber", "not set")
''    If gstrVat = "not set" Then
''        Load frmMain
''        MsgBox "The VAT Number has not yet been set." & vbCrLf & "Please make the selection in the next window.", vbInformation, "VAT Number"
''        gstrVat = frmMain.SetVAT
''        If gstrVat = "not set" Then
''            MsgBox "The VAT Number has not been changed.", vbInformation, "VAT Number"
''        End If
''    End If
''
''    'Get Address from registry
''    gstrAddress1 = GetSetting("Cash&BuyBack", "Settings", "Address1", "not set")
''    If gstrAddress1 = "not set" Then
''        MsgBox "The Address has not yet been set." & vbCrLf & "Please make the selection in the next window.", vbInformation, "Address"
''        Load frmAddress
''        frmAddress.Show vbModal
''        If gstrAddress1 = "not set" Then
''            MsgBox "The Address has not been changed.", vbInformation, "Address"
''        End If
''    Else
''        gstrAddress2 = GetSetting("Cash&BuyBack", "Settings", "Address2", "")
''        gstrAddress3 = GetSetting("Cash&BuyBack", "Settings", "Address3", "")
''    End If
''    gstrAddress = Trim(gstrAddress1) & " " & Trim(gstrAddress2) & " " & Trim(gstrAddress3)
    
    sTemplatePath = App.Path & "\templates\"
    
    frmMain.Show
    Exit Sub
    
ErrorHandle:
    Select Case Err.Number
        Case -2147467259
            MsgBox "Path to database invalid.  Specify the correct path.", vbCritical, "Invalid Path"
            Resume SetDBPath
        Case Else
            MsgBox "An error resulted in the application to terminate.", vbCritical, "Application Shutdown"
            End
    End Select
End Sub

Public Sub MissingInfo(strField As String)
    MsgBox "Please specify " & strField & ".", vbInformation, "Missing Information"
End Sub


Public Sub SetCurr(txtBox As TextBox)
    txtBox.Text = Format(txtBox.Text, "#,###.00")
End Sub
