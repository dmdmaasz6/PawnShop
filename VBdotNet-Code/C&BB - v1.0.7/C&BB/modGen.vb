Module modGen
    Public p_DBLocation As String
    Public p_CashRegister As String
    Public p_Branch As String
    Public p_AddressWH As String
    Public p_AddressWB As String
    Public p_TelWH As String
    Public p_TelWB As String
    Public p_Currency As String
    Public p_Date As Date
    Public p_Audit As Boolean
    Public p_Search_PSID As String
    Public p_Search_NSID As String
    Public p_AgreementPath As String
    Public arrItems(9, 7)  'ItemId, Quatity, Description, Serialnr, Price, SubTotal, SelDate
    Public iSaleItemCount As Integer
    Public sSaleTotal As String
    Public sAgreeTime As String

    Public Sub SetAddresses()
        p_AddressWB = "P.O. Box 72" & vbNewLine & _
                      "Walvis Bay"
        p_AddressWH = "P.O. Box 6183" & vbNewLine & _
                      "Ausspannplatz" & vbNewLine & _
                      "Windhoek"
        p_TelWB ="Tel: 064-204555"
        p_TelWH = "Tel: 061-240303 (Ausspannplatz)" & vbNewLine & _
                  "Tel: 061-301112 (Central)"

    End Sub

    Public Function ReadRegistryDB() As String
        p_DBLocation = GetSetting("Cash&BuyBack.NET", "Settings", "DBLocation", "")
        ReadRegistryDB = p_DBLocation
    End Function

    Public Function ReadRegistryAgreePath() As String
        p_AgreementPath = GetSetting("Cash&BuyBack.NET", "Settings", "AgreePath", "")
        ReadRegistryAgreePath = p_AgreementPath
    End Function

    Public Function ReadRegistryCR() As String
        p_CashRegister = GetSetting("Cash&BuyBack.NET", "Settings", "CRLocation", "")
        ReadRegistryCR = p_CashRegister
    End Function

    Public Function ReadRegistryBrnch() As String
        p_Branch = GetSetting("Cash&BuyBack.NET", "Settings", "Shop", "")
        readRegistryBrnch = p_Branch
    End Function

    Public Function ReadRegistryDate() As Date
        p_Date = GetSetting("Cash&BuyBack.NET", "Settings", "CurrDate", "")
        ReadRegistryDate = p_Date

    End Function

    Public Function ReadRegistryAudit() As Boolean

        If UCase(GetSetting("Cash&BuyBack.NET", "Settings", "Audit", "")) = "Y" Then
            p_Audit = True
        Else
            p_Audit = False
        End If

        ReadRegistryAudit = p_Audit

    End Function

    Public Sub WriteRegistryAudit(newDate As String)
        SaveSetting("Cash&BuyBack.NET", "Settings", "Audit", newDate)

    End Sub

    Public Sub WriteRegistryDate(newDate As String)
        SaveSetting("Cash&BuyBack.NET", "Settings", "CurrDate", newDate)

    End Sub

    Public Sub WriteRegistryDB(newDBLocation As String)
        SaveSetting("Cash&BuyBack.NET", "Settings", "DBLocation", newDBLocation)
    End Sub

    Public Sub WriteRegistryAgreePath(newAgreePathLocation As String)
        SaveSetting("Cash&BuyBack.NET", "Settings", "AgreePath", newAgreePathLocation)
    End Sub

    Public Sub WriteRegistryCR(newCashRegister)
        SaveSetting("Cash&BuyBack.NET", "Settings", "CRLocation", newCashRegister)
    End Sub

    Public Sub WriteRegistryBrnch(newBranch As String)
        SaveSetting("Cash&BuyBack.NET", "Settings", "Shop", newBranch)
    End Sub

    Public Sub SetCurr(txtBox As TextBox)
        txtBox.Text = FormatNumber(txtBox.Text, 2)
    End Sub

    Public Sub MissingInfo(strField As String)
        MsgBox("Please specify " & strField & ".", vbInformation, "Missing Information")
    End Sub

End Module
